using System;
using System.Runtime.InteropServices;

namespace TypeRight.Addin.UI
{
    /// <summary>Word 窗口句柄获取（老 Word 兼容）：
    /// 优先对象模型 Window.Hwnd（Word 2013+），老版本抛 COMException 时退回
    /// FindWindow("OpusApp")——OpusApp 是 Word 主窗口类名，各版本稳定不变。</summary>
    internal static class WordWindow
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowW(string className, string windowName);

        /// <summary>拿 Word 活动窗口句柄；多窗口老版本下兜底拿顶层主窗口。拿不到返回 IntPtr.Zero。</summary>
        public static IntPtr Get(NetOffice.WordApi.Application app)
        {
            try
            {
                int hwnd = app.ActiveWindow.Hwnd;
                if (hwnd != 0) return new IntPtr(hwnd);
            }
            catch { }
            return FindWindowW("OpusApp", null);
        }
    }
}
