using System;
using System.Reflection;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using TypeRight.Addin.Controllers;

namespace TypeRight.Addin
{
    /// <summary>
    /// TypeRight COM Add-in（NetOffice，免 VSTO 运行时）。
    /// 基类 NetOffice.WordApi.Tools.COMAddin 在 OnConnection 时自动把宿主包装为类型安全的
    /// Word.Application（this.Application），事件参数里收到的已是包装对象，无需也不应再手工包装。
    /// 注册：tools\install.ps1（管理员 RegAsm 或非管理员 HKCU 手工键）。
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
        public NetOffice.WordApi.Application WordApplication
        {
            get { return Application; }
        }

        public FormattingController Controller { get; private set; }

        /// <summary>任务窗格集合（基类 protected，供控制器经此访问）。</summary>
        internal NetOffice.OfficeApi.Tools.CustomTaskPaneCollection TaskPaneCollection
        {
            get { return TaskPanes; }
        }

        public TypeRightAddin()
        {
            Log("ctor");
            Controller = new FormattingController(this);

            // 任务窗格定义需在 CTPFactoryAvailable 之前注册
            var paneInfo = TaskPanes.Add(typeof(Panes.PreviewPaneHost), "智能排版");
            paneInfo.Arguments = new object[] { this };

            OnConnection += Addin_OnConnection;
            OnDisconnection += Addin_OnDisconnection;
        }

        private void Addin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            // 事件参数 application 已是 NetOffice 包装对象；类型安全的实例用基类属性 this.Application
            Log("OnConnection mode=" + connectMode + " app=" + (Application != null ? "bound" : "null"));
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
                System.Windows.Forms.MessageBox.Show("TypeRight：加载功能区失败 - " + ex.Message);
                return null;
            }
        }

        // ---- Ribbon 回调 ----

        public void OnRibbonLoad(NetOffice.OfficeApi.IRibbonUI ribbonUi)
        {
            // 预留：按需 invalidate
        }

        public void OnFormatAll(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.Run(false);
        }

        public void OnFormatSelection(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.Run(true);
        }

        public void OnProfiles(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.OpenProfileEditor();
        }

        public void OnSettings(NetOffice.OfficeApi.IRibbonControl control)
        {
            Controller.OpenSettings();
        }

}
}
