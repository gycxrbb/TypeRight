// ============================================================================
// MiniRibbon —— Office 20527 RibbonX 静默拒绝问题的判别加载项（spike，非产品代码）
// 手工 ComImport IDTExtensibility2 + IRibbonExtensibility，零外部依赖（无 NetOffice）。
// 日志：%LOCALAPPDATA%\TypeRight\mini-addin-log.txt
// 判据：Word 拿走 XML 后是否回调 OnLoad（= 功能区真的被实例化）。
// ============================================================================
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace MiniRibbon
{
    // IDTExtensibility2（extensibility.dll 的标准 GUID，手工声明免引用）
    [ComImport]
    [Guid("B65AD801-75AF-4D3D-9A31-AF2F9BC7D1B2")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst, [In] ref Array custom);
        void OnDisconnection([In] int disconnectMode, [In] ref Array custom);
        void OnAddInsUpdate([In] ref Array custom);
        void OnStartupComplete([In] ref Array custom);
        void OnBeginShutdown([In] ref Array custom);
    }

    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }

    [ComVisible(true)]
    [Guid("7A1C2E09-52B4-4C8A-9D3E-1F6B8A5C2D47")]
    [ProgId("MiniRibbon.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class MiniAddin : IDTExtensibility2, IRibbonExtensibility
    {
        private static readonly string LogFile = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TypeRight", "mini-addin-log.txt");

        private static void Log(string msg)
        {
            try
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(LogFile));
                System.IO.File.AppendAllText(LogFile,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
        }

        public MiniAddin()
        { Log("ctor —— 类已被创建"); }

        private const string RibbonXml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnLoad\">" +
            "<ribbon><tabs><tab id=\"tabMini\" label=\"Mini判别\">" +
            "<group id=\"grpMini\" label=\"判别\">" +
            "<button id=\"btnMini\" size=\"large\" label=\"点我\" onAction=\"OnBtn\"/>" +
            "</group></tab></tabs></ribbon></customUI>";

        // ── IDTExtensibility2 ──
        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        { Log("OnConnection mode=" + connectMode); }
        public void OnDisconnection(int disconnectMode, ref Array custom)
        { Log("OnDisconnection mode=" + disconnectMode); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { Log("OnStartupComplete"); }
        public void OnBeginShutdown(ref Array custom) { }

        // ── IRibbonExtensibility ──
        public string GetCustomUI(string ribbonID)
        {
            Log("GetCustomUI(" + ribbonID + ") -> " + RibbonXml.Length + " 字符");
            return RibbonXml;
        }

        // ── RibbonX 回调（IDispatch 名字解析）──
        public void OnLoad(object ribbonUi)
        { Log("★ Ribbon onLoad —— 功能区被实例化（判据达成）"); }
        public void OnBtn(object control)
        { Log("按钮被点击"); }
    }
}
