using System;
using System.Collections.Generic;
using System.Linq;
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
    /// 编排器：扫描 → 角色判定 →（可选）LLM 低置信增强 → 预览任务窗格 → 应用 → 统计反馈。
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

        public FormattingController(TypeRightAddin addin)
        {
            if (addin == null) throw new ArgumentNullException("addin");
            _addin = addin;
        }

        public Application App { get { return _addin.WordApplication; } }

        public AppSettings Settings { get { return _settings; } }

        // ---- 任务窗格 ----

        public void AttachHost(Panes.PreviewPaneHost host)
        {
            _paneHost = host;
        }

        public void Shutdown()
        {
            try { if (_fallbackWindow != null) _fallbackWindow.Close(); } catch { }
        }

        /// <summary>显示预览窗格：优先 CTP 任务窗格，创建失败回落为无模式窗口。</summary>
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

        private void ShowFallbackWindow()
        {
            lock (_uiLock)
            {
                if (_fallbackWindow == null || !_fallbackWindow.IsLoaded)
                {
                    _fallbackWindow = new System.Windows.Window
                    {
                        Title = "TypeRight 智能排版预览",
                        Width = 520,
                        Height = 640,
                        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                        Content = new Panes.PreviewControl(),
                        ShowInTaskbar = false
                    };
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

        /// <summary>隐藏预览窗格。</summary>
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

        public void Run(bool selectionOnly)
        {
            try
            {
                Document doc = null;
                try { doc = App.ActiveDocument; }
                catch { }
                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show("请先打开一个文档。", "TypeRight");
                    return;
                }

                Range restrict = null;
                if (selectionOnly)
                {
                    try
                    {
                        int selStart = App.Selection.Range.Start;
                        int selEnd = App.Selection.Range.End;
                        if (selEnd > selStart)
                        {
                            restrict = doc.Range(selStart, selEnd);
                        }
                    }
                    catch { }
                    if (restrict == null)
                    {
                        System.Windows.Forms.MessageBox.Show("当前没有选区，将排版整个文档。", "TypeRight");
                    }
                }

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

                // LLM 增强（可选，仅低置信段落；失败回落规则）
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
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("扫描失败：" + ex.Message, "TypeRight");
            }
        }

        private void NotifyPreview(ScanResult scan, bool selectionMode, string note, Range restrict)
        {
            Panes.PreviewControl target = null;
            if (_paneHost != null) target = _paneHost.PreviewControl;
            if (target == null && _fallbackWindow != null)
                target = _fallbackWindow.Content as Panes.PreviewControl;
            if (target != null)
            {
                target.Bind(this);
                target.ViewModel.LoadScan(this, scan, selectionMode, note, restrict);
            }
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

        // ---- 应用排版（由预览界面触发） ----

        public void Apply(IDictionary<int, ParagraphRole> assignments, Profile profile, bool mountStyles, bool convertFormulas, Range selection)
        {
            Document doc = null;
            try { doc = App.ActiveDocument; }
            catch { }
            if (doc == null)
            {
                System.Windows.Forms.MessageBox.Show("文档已关闭，无法应用。", "TypeRight");
                return;
            }

            var request = new FormatRequest
            {
                Profile = profile,
                RoleAssignments = assignments,
                MountStyles = mountStyles,
                ConvertFormulas = convertFormulas,
                Selection = selection
            };

            ApplyStats stats;
            var oldCursor = System.Windows.Forms.Cursor.Current;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            try
            {
                stats = new FormatApplier().Apply(App, doc, request);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("排版已整体回退：" + ex.Message, "TypeRight");
                return;
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = oldCursor;
            }

            System.Windows.Forms.MessageBox.Show(
                "排版完成：\r\n" + stats.Summary() +
                "\r\n\r\n一次 Ctrl+Z 可完全撤销本次操作。", "TypeRight");
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
