using System;
using System.Reflection;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using TypeRight.Addin.Controllers;

namespace TypeRight.Addin
{
    /// <summary>
    /// TypeRight COM Add-in（PRD v2：指认式批量排版）。
    /// UI 只有：Ribbon 3 按钮 + 右键菜单 + 浮动格式条 + 轻提示。无任务窗格、无预览弹窗。
    /// </summary>
    [Guid("6C7C9713-51A1-4A55-8C0D-3A4E5F607182")]
    [ProgId("TypeRight.Connect")]
    [ComVisible(true)]
    public sealed class TypeRightAddin : NetOffice.WordApi.Tools.COMAddin
    {
        private static readonly string LogFile = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TypeRight", "addin-log.txt");

        /// <summary>生命周期诊断日志（%LOCALAPPDATA%\TypeRight\addin-log.txt）。</summary>
        public static void Log(string message)
        {
            try
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(LogFile));
                System.IO.File.AppendAllText(LogFile,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + message + Environment.NewLine,
                    System.Text.Encoding.UTF8);
            }
            catch { }
        }

        /// <summary>资源管理器中定位诊断日志（帮助窗入口）。</summary>
        public static void OpenLogFolder()
        {
            try
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(LogFile));
                System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + LogFile + "\"");
            }
            catch (Exception ex) { Log("OpenLogFolder: " + ex.Message); }
        }

        /// <summary>一键诊断（发给作者排障）：版本 + 系统 + 最近日志尾部。</summary>
        public static string CollectDiagnostics(NetOffice.WordApi.Application app)
        {
            var sb = new System.Text.StringBuilder();
            try
            {
                sb.AppendLine("TypeRight v" + typeof(TypeRightAddin).Assembly.GetName().Version.ToString(3));
                sb.AppendLine("OS: " + Environment.OSVersion.VersionString + "  " + (Environment.Is64BitProcess ? "x64" : "x86") + " 进程");
                try { sb.AppendLine("Word: " + app.Version); } catch { }
            }
            catch { }
            sb.AppendLine("── addin-log 尾部 ──");
            try
            {
                var lines = System.IO.File.ReadAllLines(LogFile, System.Text.Encoding.UTF8);
                for (int i = Math.Max(0, lines.Length - 80); i < lines.Length; i++)
                    sb.AppendLine(lines[i]);
            }
            catch (Exception ex) { sb.AppendLine("(日志读取失败: " + ex.Message + ")"); }
            return sb.ToString();
        }

        /// <summary>首次运行提示标记（一次即不再弹）。</summary>
        private static string FirstRunFlagFile
        {
            get
            {
                return System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TypeRight", "first-run-shown.txt");
            }
        }

        /// <summary>首次运行提示（发布支持）：三步用法，展示一次后落标记文件。</summary>
        public void ShowFirstRunTipIfDue()
        {
            try
            {
                if (System.IO.File.Exists(FirstRunFlagFile)) return;
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FirstRunFlagFile));
                System.IO.File.WriteAllText(FirstRunFlagFile, DateTime.Now.ToString("s"), System.Text.Encoding.UTF8);
                Controller.ShowToast("TypeRight 已就绪：光标点进段落 →「同类选择」→ 浮动条批量改格式。详细用法见选项卡里的「使用说明」。", 9000);
                Log("FirstRunTip: shown");
            }
            catch (Exception ex) { Log("FirstRunTip: " + ex.Message); }
        }

        /// <summary>宿主 Word 应用（基类在 OnConnection 时已包装为类型安全实例）。</summary>
        public NetOffice.WordApi.Application WordApplication { get { return Application; } }

        public TypeRightController Controller { get; private set; }

        /// <summary>功能区对象（主题变化时 Invalidate 刷新自定义图标配色）。</summary>
        private NetOffice.OfficeApi.IRibbonUI _ribbon;
        private UI.OfficeThemeKind? _iconTheme;

        public TypeRightAddin()
        {
            Log("ctor");
            Controller = new TypeRightController(this);

            OnConnection += Addin_OnConnection;
            OnDisconnection += Addin_OnDisconnection;
        }

        private void Addin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Log("OnConnection mode=" + connectMode + " app=" + (Application != null ? "bound" : "null"));
            try
            {
                // 点击文档其他处 → 关闭浮动条（PRD P0-2）
                Application.WindowSelectionChangeEvent += Application_WindowSelectionChange;
                // 关闭文档前清掉同类选择的临时高亮（否则用户保存文档会把黄色标记写进文件）
                Application.DocumentBeforeCloseEvent += Application_DocumentBeforeClose;
            }
            catch (Exception ex)
            {
                Log("事件订阅失败: " + ex.Message);
            }
            try
            {
                // 后台预热字体枚举（InstalledFontCollection 要 1~2s，别落在首次点击路径上）
                System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                {
                    try { UI.FontListProvider.GetFontList(); } catch { }
                });
            }
            catch { }
            try
            {
                // 暴露给自动化验证脚本（COMAddIns.Item("TypeRight.Connect").Object）
                var comAddin = addInInst as NetOffice.OfficeApi.COMAddIn;
                if (comAddin != null)
                {
                    comAddin.Object = this;
                    Log("Object set via wrapper");
                }
                else if (addInInst != null)
                {
                    addInInst.GetType().InvokeMember("Object",
                        System.Reflection.BindingFlags.SetProperty, null, addInInst, new object[] { this });
                    Log("Object set via IDispatch");
                }
                else
                {
                    Log("Object set skipped: addInInst null");
                }
            }
            catch (Exception ex)
            {
                Log("COMAddIn.Object 设置失败: " + ex.Message);
            }

            Log("Version: " + typeof(TypeRightAddin).Assembly.GetName().Version.ToString(3));
            try
            {
                int major = WordCompat.MajorVersion(Application);
                Log("Host Word: " + Application.Version + (major == 14 ? "（2010 兼容模式：无自定义撤销）" : major == 15 ? "（2013）" : ""));
            }
            catch { }
        }

        private void Application_WindowSelectionChange(NetOffice.WordApi.Selection selection)
        {
            if (!_firstRunChecked)
            {
                _firstRunChecked = true;
                // 首次运行提示挂在用户第一次操作时弹（OnConnection 时机太早，
                // Word 窗口尚未建好，SetTimer 队列会拿不到 hwnd）
                ShowFirstRunTipIfDue();
            }
            if (!_hostCompatNoted)
            {
                _hostCompatNoted = true;
                // Word 2010：UndoRecord 不存在 → 批量操作的撤销退化为多条记录，主动告知一次
                try
                {
                    if (WordCompat.MajorVersion(Application) == 14)
                        Controller.ShowToast("Word 2010 兼容模式：功能全部可用；仅撤销粒度降级（一次批量修改需多按几次 Ctrl+Z 才能完全回退）。", 8000);
                }
                catch { }
            }
            if (_prewarmQueued == false)
            {
                _prewarmQueued = true;
                // 空闲期预建浮条（隐藏）：WPF 初始化 ~1s 挪出"同类选择"首次点击
                QueueOnWordUiThread(() => Controller.PrewarmFloatBar(), 2000);
            }
            if (_pendingAutomation != null)
            {
                RunPendingAutomation();
                return;
            }
            try
            {
                Controller.HandleSelectionChange(selection);
            }
            catch (Exception ex)
            {
                Log("HandleSelectionChange: " + ex.Message);
            }
        }

        private void Application_DocumentBeforeClose(NetOffice.WordApi.Document doc, ref bool cancel)
        {
            try
            {
                Controller.ClearSessionIfDoc(doc);
            }
            catch (Exception ex)
            {
                Log("DocumentBeforeClose: " + ex.Message);
            }
        }

        private void Addin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Log("OnDisconnection mode=" + removeMode);
            Controller.Shutdown();
        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Log("OnError " + methodKind + ": " + exception);
            System.Windows.Forms.MessageBox.Show("TypeRight 出错（" + methodKind + "）：\r\n" + exception.Message);
        }

        public override string GetCustomUI(string ribbonId)
        {
            try
            {
                Assembly asm = typeof(TypeRightAddin).Assembly;
                using (var stream = asm.GetManifestResourceStream("TypeRight.Addin.Ribbon.xml"))
                using (var reader = new System.IO.StreamReader(stream, System.Text.Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Log("加载功能区失败: " + ex.Message);
                return null;
            }
        }

        // ---- Ribbon 回调 ----

        public void OnRibbonLoad(NetOffice.OfficeApi.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        /// <summary>自定义图标（getImage）：按宿主 Word 主题出深/浅两套配色。</summary>
        public object OnGetImage(NetOffice.OfficeApi.IRibbonControl control)
        {
            try
            {
                UI.OfficeThemeKind theme = UI.OfficeTheme.Detect(Application);
                _iconTheme = theme;
                return UI.RibbonIcons.GetImage(control.Id, theme);
            }
            catch (Exception ex)
            {
                Log("OnGetImage[" + control.Id + "]: " + ex.Message);
                return null;
            }
        }

        /// <summary>宿主主题变化时刷新功能区图标配色（悬浮条每次打开时调用）。</summary>
        public void RefreshRibbonTheme()
        {
            try
            {
                UI.OfficeThemeKind now = UI.OfficeTheme.Detect(Application, force: true);
                if (_iconTheme.HasValue && _iconTheme.Value != now && _ribbon != null)
                    _ribbon.Invalidate();
            }
            catch { }
        }

        public void OnSimilar(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.SimilarSelect();
        }

        public void OnFormulas(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.ConvertFormulas();
        }

        public void OnProfiles(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.OpenProfiles();
        }

        public void OnProfilesSave(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.OpenProfiles(saveMode: true);
        }

        public void OnProfilesManage(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.OpenProfiles(saveMode: false);
        }

        public void OnSettings(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.OpenSettings();
        }

        public void OnHelp(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.ShowHelp();
        }

        // ── 自动化验证钩子（测试脚本经 COMAddIn.Object 调用，不在用户界面暴露）──
        // 跨进程 COM 调用落在 RPC 线程；实测该线程与事件回调线程上 Word COM 均不可靠
        // （Range.Information(wdWithInTable) 返回假值）。用 SetTimer 挂到 Word 窗口：
        // WM_TIMER 由 Word UI 线程的消息循环分发，回调语境与 Ribbon 点击一致（可靠）。

        private Action _pendingAutomation;
        private Win32TimerProc _timerProc;   // 持引用防 GC 回收原生回调桩
        private bool _prewarmQueued;         // 浮条空闲预热只排一次
        private bool _firstRunChecked;       // 首次运行提示只查一次
        private bool _hostCompatNoted;       // 老版本兼容提示只弹一次

        private delegate void Win32TimerProc(IntPtr hWnd, uint msg, IntPtr idEvent, uint dwTime);

        [DllImport("user32.dll")] private static extern IntPtr SetTimer(IntPtr hWnd, IntPtr idEvent, uint elapse, Win32TimerProc proc);
        [DllImport("user32.dll")] private static extern bool KillTimer(IntPtr hWnd, IntPtr idEvent);

        private static readonly IntPtr AutomationTimerId = new IntPtr(0x54525);

        public void AutomationSimilarSelect()
        {
            QueueOnWordUiThread(Controller.SimilarSelect);
        }

        public void AutomationCloseSession()
        {
            QueueOnWordUiThread(() => Controller.CloseSession(false));
        }

        /// <summary>强制悬浮条主题："dark" / "darkgray" / "light" / 其他=自动检测（仅 UI，不动文档）。</summary>
        public void AutomationForceFloatTheme(string theme)
        {
            QueueOnWordUiThread(() => Controller.ForceFloatBarTheme(theme));
        }

        /// <summary>自动化专用：以 ShowActivated=false 重新显示悬浮条（后台进程抢不到前台时
        /// Activate 会触发 Deactivated→Hide；不激活显示则保持可见，供 PrintWindow 取证）。</summary>
        public void AutomationShowFloatBar()
        {
            QueueOnWordUiThread(() =>
            {
                try
                {
                    Controller.ShowFloatBarForAutomation();
                }
                catch (Exception ex) { Log("AutomationShowFloatBar: " + ex.Message); }
            });
        }

        private void QueueOnWordUiThread(Action action, uint delayMs = 10)
        {
            _pendingAutomation = action;
            try
            {
                IntPtr hwnd = UI.WordWindow.Get(Application);   // 2013+ 走对象模型；老 Word 走 FindWindow 兜底
                if (hwnd == IntPtr.Zero) { Log("queue: no word hwnd"); return; }
                _timerProc = OnAutomationTimer;
                SetTimer(hwnd, AutomationTimerId, delayMs, _timerProc);
            }
            catch (Exception ex)
            {
                Log("queue SetTimer failed: " + ex.Message);
            }
        }

        private void OnAutomationTimer(IntPtr hWnd, uint msg, IntPtr idEvent, uint dwTime)
        {
            try { KillTimer(hWnd, idEvent); } catch { }
            Action action = System.Threading.Interlocked.Exchange(ref _pendingAutomation, null);
            if (action == null) return;
            Log("automation timer fire (ui thread)");
            try { action(); }
            catch (Exception ex) { Log("automation action: " + ex); }
        }

        private void RunPendingAutomation()
        {
            Action pending = System.Threading.Interlocked.Exchange(ref _pendingAutomation, null);
            if (pending == null) return;
            try { pending(); }
            catch (Exception ex) { Log("pending automation: " + ex); }
        }
    }
}
