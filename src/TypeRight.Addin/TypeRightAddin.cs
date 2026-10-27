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
        }

        private void Application_WindowSelectionChange(NetOffice.WordApi.Selection selection)
        {
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

        public void OnSettings(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.OpenSettings();
        }
    }
}
