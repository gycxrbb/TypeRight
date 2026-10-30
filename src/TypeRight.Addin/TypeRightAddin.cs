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
            }
            catch (Exception ex)
            {
                Log("WindowSelectionChange 订阅失败: " + ex.Message);
            }
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
        }

        private void Application_WindowSelectionChange(NetOffice.WordApi.Selection selection)
        {
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

        // ── 自动化验证钩子（测试脚本经 COMAddIn.Object 调用，不在用户界面暴露）──
        // 跨进程 COM 调用落在 RPC 线程；实测该线程与事件回调线程上 Word COM 均不可靠
        // （Range.Information(wdWithInTable) 返回假值）。用 SetTimer 挂到 Word 窗口：
        // WM_TIMER 由 Word UI 线程的消息循环分发，回调语境与 Ribbon 点击一致（可靠）。

        private Action _pendingAutomation;
        private Win32TimerProc _timerProc;   // 持引用防 GC 回收原生回调桩

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

        private void QueueOnWordUiThread(Action action)
        {
            _pendingAutomation = action;
            try
            {
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = new IntPtr(Application.ActiveWindow.Hwnd); } catch { }
                if (hwnd == IntPtr.Zero) { Log("queue: no word hwnd"); return; }
                _timerProc = OnAutomationTimer;
                SetTimer(hwnd, AutomationTimerId, 10, _timerProc);
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
