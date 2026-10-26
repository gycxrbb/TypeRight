using System;
using System.Runtime.InteropServices;
using NetOffice.OfficeApi.Enums;
using NetOffice.OfficeApi.Tools;

namespace TypeRight.Addin.Panes
{
    /// <summary>
    /// 预览任务窗格宿主（WinForms UserControl）。
    /// NetOffice 以 Type.FullName 作为 ProgId 调用 ICTPFactory.CreateCTP，
    /// 因此 COM 注册必须提供 ProgId = TypeRight.Addin.Panes.PreviewPaneHost。
    /// 内部用 ElementHost 承载 WPF 预览控件。
    /// </summary>
    [ComVisible(true)]
    [Guid("B3E4D2A1-9F56-4C71-8D2E-6A7B8C9D0E1F")]
    [ProgId("TypeRight.Addin.Panes.PreviewPaneHost")]
    public partial class PreviewPaneHost :
        System.Windows.Forms.UserControl,
        ITaskPane,
        IObjectSafety
    {
        private const int SafeForCaller = 0x00000001;
        private const int SafeForData = 0x00000002;

        // ---- IObjectSafety ----

        int IObjectSafety.GetInterfaceSafetyOptions(ref Guid riid, out int supported, out int enabled)
        {
            supported = SafeForCaller | SafeForData;
            enabled = SafeForCaller | SafeForData;
            return 0;
        }

        int IObjectSafety.SetInterfaceSafetyOptions(ref Guid riid, int mask, int enabled)
        {
            return 0;
        }
        private System.Windows.Forms.Integration.ElementHost _elementHost;

        public PreviewControl PreviewControl { get; private set; }

        public PreviewPaneHost()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            _elementHost = new System.Windows.Forms.Integration.ElementHost
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                Child = (PreviewControl = new PreviewControl())
            };
            SuspendLayout();
            Controls.Add(_elementHost);
            Name = "PreviewPaneHost";
            ResumeLayout(false);
        }

        // ---- ITaskPane：由 NetOffice 在 CTP 创建后回调 ----

        public void OnConnection(NetOffice.ICOMObject application, NetOffice.OfficeApi._CustomTaskPane taskPane, object[] arguments)
        {
            TypeRightAddin.Log("CTP OnConnection (任务窗格创建成功)");
            var addin = arguments != null && arguments.Length > 0 ? arguments[0] as TypeRightAddin : null;
            if (addin == null || addin.Controller == null) return;
            addin.Controller.AttachHost(this);
            PreviewControl.Bind(addin.Controller);
            try { if (taskPane != null) taskPane.Width = 420; }
            catch { }
        }

        // ---- ITaskPane ----

        public void OnDisconnection()
        {
        }

        public void OnDockPositionChanged(MsoCTPDockPosition position)
        {
        }

        public void OnVisibleStateChanged(bool visible)
        {
        }
    }

    /// <summary>Office CTP 宿主要求 ActiveX 控件标记脚本安全。</summary>
    [ComImport]
    [Guid("CB5BDC81-93C1-11cf-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectSafety
    {
        [PreserveSig]
        int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions);

        [PreserveSig]
        int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions);
    }
}
