using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice.WordApi;
using TypeRight.Classification;
using TypeRight.Formatting;
using TypeRight.Llm;
using TypeRight.Profiles;
using TypeRight.Roles;
using TypeRight.Scanning;

namespace TypeRight.Addin.Controllers
{
    /// <summary>
    /// 编排器：扫描 → 角色判定 →（可选）LLM 低置信增强 → 预览面板 → 应用 → 结果反馈。
    /// 所有 Word 互操作都在 Word 主线程（Ribbon 回调线程）执行。
    /// </summary>
    public sealed class FormattingController
    {
        private readonly TypeRightAddin _addin;
        private readonly ProfileStore _profileStore = new ProfileStore();

        private AppSettings _settings = SettingsStore.Load();
        private List<LoadedProfile> _profiles = new List<LoadedProfile>();
        private List<string> _profileErrors = new List<string>();

        private Panes.PreviewPaneHost _paneHost;
        private System.Windows.Window _fallbackWindow;
        private object _uiLock = new object();

        /// <summary>最近一次扫描的文档全名（应用前校验，防止用户切换文档后误应用）。</summary>
        private string _scannedDocFullName;

        public FormattingController(TypeRightAddin addin)
        {
            if (addin == null) throw new ArgumentNullException("addin");
            _addin = addin;
        }

        public Application App { get { return _addin.WordApplication; } }

        public AppSettings Settings { get { return _settings; } }

        // ---- 任务窗格 / 面板 ----

        public void AttachHost(Panes.PreviewPaneHost host)
        {
            _paneHost = host;
        }

        public void Shutdown()
        {
            try { if (_fallbackWindow != null) _fallbackWindow.Close(); } catch { }
        }

        /// <summary>预览面板是否可用（CTP 任务窗格已创建，或浮动窗口已打开）。</summary>
        public bool IsPaneSurfaceReady()
        {
            return _paneHost != null || (_fallbackWindow != null && _fallbackWindow.IsLoaded);
        }

        /// <summary>显示预览面板：优先 CTP 停靠任务窗格，未创建则回落为贴靠 Word 右侧的浮动面板。</summary>
        public void ShowPreviewPane()
        {
            if (_paneHost != null)
            {
                try
                {
                    var panes = _addin.TaskPaneCollection;
                    if (panes != null)
                    {
                        foreach (var info in panes)
                        {
                            if (info == null) continue;
                            try { info.Visible = true; return; }
                            catch { }
                        }
                    }
                }
                catch { }
            }
            ShowFallbackWindow();
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        private void ShowFallbackWindow()
        {
            lock (_uiLock)
            {
                if (_fallbackWindow == null || !_fallbackWindow.IsLoaded)
                {
                    _fallbackWindow = new System.Windows.Window
                    {
                        Title = "TypeRight 智能排版",
                        Width = 470,
                        Height = 760,
                        WindowStartupLocation = System.Windows.WindowStartupLocation.Manual,
                        ResizeMode = System.Windows.ResizeMode.CanResize,
                        ShowInTaskbar = false,
                        Content = new Panes.PreviewControl()
                    };

                    // 贴靠 Word 窗口右侧并成为其所属窗口（保持在 Word 之上）
                    try
                    {
                        IntPtr wordHwnd = new IntPtr(App.ActiveWindow.Hwnd);
                        RECT r;
                        if (GetWindowRect(wordHwnd, out r))
                        {
                            uint dpi = GetDpiForWindow(wordHwnd);
                            double scale = dpi > 0 ? dpi / 96.0 : 1.0; // 物理像素 → WPF DIP
                            _fallbackWindow.Left = r.Right / scale - _fallbackWindow.Width - 16;
                            _fallbackWindow.Top = r.Top / scale + 96;
                        }
                        new System.Windows.Interop.WindowInteropHelper(_fallbackWindow).Owner = wordHwnd;
                    }
                    catch { }

                    _fallbackWindow.Closed += (s, e) => _fallbackWindow = null;
                    _fallbackWindow.Show();
                    if (_fallbackWindow.Content is Panes.PreviewControl ctrl)
                        ctrl.Bind(this);
                }
                else
                {
                    _fallbackWindow.Activate();
                }
            }
        }

        /// <summary>隐藏预览面板。</summary>
        public void HidePreviewPane()
        {
            try
            {
                if (_paneHost != null)
                {
                    foreach (var info in _addin.TaskPaneCollection)
                    {
                        if (info != null) { info.Visible = false; return; }
                    }
                }
                if (_fallbackWindow != null) _fallbackWindow.Close();
            }
            catch { }
        }

        // ---- 主流程 ----

        /// <summary>扫描当前文档并打开预览面板。</summary>
        public void Run(bool selectionOnly)
        {
            try
            {
                Document doc = GetActiveDocument();
                if (doc == null)
                {
                    // 无文档：打开面板到欢迎页（可生成示例文档体验）
                    ReloadProfiles();
                    ShowPreviewPane();
                    NotifyWelcome();
                    return;
                }

                Range restrict = null;
                if (selectionOnly)
                {
                    try
                    {
                        int selStart = App.Selection.Range.Start;
                        int selEnd = App.Selection.Range.End;
                        if (selEnd > selStart) restrict = doc.Range(selStart, selEnd);
                    }
                    catch { }
                }

                ScanAndNotify(doc, restrict);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("扫描失败：" + ex.Message, "TypeRight");
            }
        }

        private Document GetActiveDocument()
        {
            try { return App.ActiveDocument; }
            catch { return null; }
        }

        private void ScanAndNotify(Document doc, Range restrict)
        {
            var watchCursor = System.Windows.Forms.Cursor.Current;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            ScanResult scan;
            try
            {
                var scanner = new DocumentScanner(App);
                scan = scanner.Scan(doc, restrict);
                RoleClassifier.Classify(scan.Features);
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = watchCursor;
            }

            try { _scannedDocFullName = doc.FullName; } catch { _scannedDocFullName = null; }

            string llmNote = null;
            if (_settings.Llm != null && _settings.Llm.Enabled)
            {
                var resolver = new AmbiguityResolver(_settings.Llm);
                var result = resolver.Resolve(scan.Features);
                if (!result.UsedLlm) llmNote = "LLM 增强失败，已回落规则结果：" + result.Error;
                else if (result.AdjustedCount > 0) llmNote = "LLM 调整了 " + result.AdjustedCount + " 个低置信段落";
            }

            ReloadProfiles();
            ShowPreviewPane();
            NotifyPreview(scan, restrict != null, llmNote, restrict);
        }

        /// <summary>生成示例文档（故意脏格式）并立即扫描，让用户零风险体验完整流程。</summary>
        public void CreateSampleDocument()
        {
            try
            {
                Document doc = SampleDocBuilder.Build(App);
                try { doc.ActiveWindow.Activate(); } catch { }
                ScanAndNotify(doc, null);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("生成示例文档失败：" + ex.Message, "TypeRight");
            }
        }

        private void NotifyPreview(ScanResult scan, bool selectionMode, string note, Range restrict)
        {
            Panes.PreviewControl target = GetPreviewTarget();
            if (target != null)
            {
                target.Bind(this);
                target.ViewModel.LoadScan(this, scan, selectionMode, note, restrict);
            }
        }

        private void NotifyWelcome()
        {
            Panes.PreviewControl target = GetPreviewTarget();
            if (target != null)
            {
                target.Bind(this);
                target.ViewModel.ShowWelcome(this);
            }
        }

        private Panes.PreviewControl GetPreviewTarget()
        {
            if (_paneHost != null) return _paneHost.PreviewControl;
            if (_fallbackWindow != null) return _fallbackWindow.Content as Panes.PreviewControl;
            return null;
        }

        public void ReloadProfiles()
        {
            var errors = new List<string>();
            _profiles = _profileStore.LoadAll(errors);
            _profileErrors = errors;
        }

        public List<LoadedProfile> Profiles { get { return _profiles; } }
        public List<string> ProfileErrors { get { return _profileErrors; } }
        public string LastProfileId { get { return _settings.LastProfileId; } }

        public void SaveLastProfile(string id)
        {
            _settings.LastProfileId = id;
            try { SettingsStore.Save(_settings); } catch { }
        }

        /// <summary>当前活动文档显示名（面板标题用）。</summary>
        public string ActiveDocumentName()
        {
            var doc = GetActiveDocument();
            if (doc == null) return null;
            try { return System.IO.Path.GetFileName(doc.FullName); }
            catch { return doc.Name; }
        }

        // ---- 应用排版（由预览界面触发）----

        public sealed class ApplyOutcome
        {
            public ApplyStats Stats;
            public string Error;   // 非空 = 失败（已整体回退）
        }

        public ApplyOutcome Apply(IDictionary<int, ParagraphRole> assignments, Profile profile, bool mountStyles, bool convertFormulas, Range selection)
        {
            Document doc = GetActiveDocument();
            if (doc == null)
                return new ApplyOutcome { Error = "文档已关闭，无法应用。" };

            // 防呆：扫描后用户切换了文档
            string currentName;
            try { currentName = doc.FullName; } catch { currentName = null; }
            if (!string.IsNullOrEmpty(_scannedDocFullName) &&
                !string.Equals(currentName, _scannedDocFullName, StringComparison.OrdinalIgnoreCase))
            {
                return new ApplyOutcome { Error = "预览结果对应的是另一个文档（" + _scannedDocFullName + "）。\r\n请切换回该文档，或重新扫描当前文档。" };
            }

            var request = new FormatRequest
            {
                Profile = profile,
                RoleAssignments = assignments,
                MountStyles = mountStyles,
                ConvertFormulas = convertFormulas,
                Selection = selection
            };

            var oldCursor = System.Windows.Forms.Cursor.Current;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            try
            {
                var stats = new FormatApplier().Apply(App, doc, request);
                return new ApplyOutcome { Stats = stats };
            }
            catch (Exception ex)
            {
                return new ApplyOutcome { Error = "排版执行失败，已整体回退：" + ex.Message };
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = oldCursor;
            }
        }

        // ---- 设置 / 方案管理 ----

        public void OpenSettings()
        {
            try
            {
                var win = new Windows.SettingsWindow(SettingsStore.Load()) { Owner = null };
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("打开设置失败：" + ex.Message, "TypeRight");
            }
            finally
            {
                _settings = SettingsStore.Load();
            }
        }

        public void OpenProfileEditor()
        {
            try
            {
                ReloadProfiles();
                var win = new Windows.ProfileEditorWindow(this);
                win.ShowDialog();
                ReloadProfiles();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("打开方案管理失败：" + ex.Message, "TypeRight");
            }
        }
    }
}
