using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using TypeRight.Addin.UI;

namespace TypeRight.Addin.UI
{
    /// <summary>
    /// 功能区自定义图标（设计稿：Office 风格双色线性图标，统一尺寸/线宽/配色）。
    /// WPF 矢量绘制 → RenderTargetBitmap(Pbgra32) → HBITMAP(32bpp ARGB) → IPictureDisp。
    /// 按宿主主题出两套颜色（深色主题用浅笔画，浅色主题用深笔画），保证深浅背景都可辨。
    /// 大按钮 32px，菜单项/右键菜单 16px；自定义位图无原生图标的多 DPI 版本，
    /// 高缩放下由 Word 拉伸（Office 自定义图标的通用限制）。
    /// </summary>
    internal static class RibbonIcons
    {
        // 大按钮：同类选择 / 公式转换 / 方案（菜单）；小图标：右键菜单 + 方案子项
        public static object GetImage(string controlId, OfficeThemeKind theme)
        {
            bool small = controlId == "ctxSimilar" || controlId == "btnProfileSave" || controlId == "btnProfileManage" || controlId == "btnHelp";
            int size = small ? 16 : 32;
            RenderTargetBitmap bmp = Render(controlId, size, theme);
            return ToPictureDisp(bmp);
        }

        // ───────── 绘制 ─────────

        private static RenderTargetBitmap Render(string id, int size, OfficeThemeKind theme)
        {
            bool dark = theme != OfficeThemeKind.Light;
            Color fg = dark ? Color.FromRgb(0xE8, 0xE8, 0xE8) : Color.FromRgb(0x44, 0x44, 0x44);
            Color acc = dark ? Color.FromRgb(0x5B, 0x9C, 0xF8) : Color.FromRgb(0x2B, 0x57, 0x9A);
            double u = size / 32.0; // 16px 图标按 32 设计稿等比缩

            var visual = new DrawingVisual();
            using (DrawingContext dc = visual.RenderOpen())
            {
                switch (id)
                {
                    case "btnSimilar":
                    case "ctxSimilar":
                        DrawSimilar(dc, u, fg, acc);
                        break;
                    case "btnFormulas":
                        DrawFormulas(dc, u, fg, acc);
                        break;
                    case "mnuProfiles":
                        DrawProfiles(dc, u, fg, acc);
                        break;
                    case "btnProfileSave":
                        DrawSave(dc, u, fg, acc);
                        break;
                    case "btnProfileManage":
                        DrawManage(dc, u, fg, acc);
                        break;
                    case "btnHelp":
                        DrawHelp(dc, u, fg, acc);
                        break;
                }
            }

            var bmp = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(visual);
            bmp.Freeze();
            return bmp;
        }

        private static Brush Solid(Color c, double opacity)
        {
            var b = new SolidColorBrush(Color.FromArgb((byte)(opacity * 255), c.R, c.G, c.B));
            b.Freeze();
            return b;
        }

        private static Pen Pen(Color c, double thickness, double opacity = 1.0)
        {
            var p = new Pen(Solid(c, opacity), thickness)
            {
                StartLineCap = PenLineCap.Round,
                EndLineCap = PenLineCap.Round,
                DashCap = PenLineCap.Round,
                LineJoin = PenLineJoin.Round
            };
            p.Freeze();
            return p;
        }

        private static void Line(DrawingContext dc, Pen p, double x1, double y1, double x2, double y2)
        {
            dc.DrawLine(p, new Point(x1, y1), new Point(x2, y2));
        }

        private static void Bar(DrawingContext dc, Brush b, double x, double y, double w, double h)
        {
            dc.DrawRoundedRectangle(b, null, new Rect(x, y, w, h), h / 2, h / 2);
        }

        /// <summary>同类选择：多行文本（灰）+ 虚线选框（蓝，圈住后两行）。</summary>
        private static void DrawSimilar(DrawingContext dc, double u, Color fg, Color acc)
        {
            dc.PushTransform(new ScaleTransform(u, u));
            // 三行"文本"
            Bar(dc, Solid(fg, 0.9), 5.0, 6.0, 14.0, 2.8);
            Bar(dc, Solid(fg, 0.9), 5.0, 13.0, 22.0, 2.8);
            Bar(dc, Solid(fg, 0.9), 5.0, 20.0, 16.0, 2.8);
            // 蓝色虚线选择框（跑马灯语义），圈住第 2~3 行
            var pen = Pen(acc, 1.7);
            pen = new Pen(Solid(acc, 1.0), 1.7)
            {
                StartLineCap = PenLineCap.Round,
                EndLineCap = PenLineCap.Round,
                DashCap = PenLineCap.Round,
                LineJoin = PenLineJoin.Round,
                DashStyle = new DashStyle(new double[] { 2.6, 1.8 }, 0)
            };
            pen.Freeze();
            dc.DrawRoundedRectangle(null, pen, new Rect(2.6, 10.0, 26.8, 16.0), 2.4, 2.4);
            dc.Pop();
        }

        /// <summary>公式转换：根号（灰）+ 双向转换箭头（蓝）。</summary>
        private static void DrawFormulas(DrawingContext dc, double u, Color fg, Color acc)
        {
            dc.PushTransform(new ScaleTransform(u, u));
            // √ 根号 + 根号下横线
            var geo = new StreamGeometry();
            using (StreamGeometryContext sgc = geo.Open())
            {
                sgc.BeginFigure(new Point(5, 15.5), false, false);
                sgc.LineTo(new Point(9.2, 8.0), true, false);
                sgc.LineTo(new Point(12.8, 23.5), true, false);
                sgc.LineTo(new Point(26, 23.5), true, false);
            }
            geo.Freeze();
            dc.DrawGeometry(null, Pen(fg, 2.6), geo);
            // 右上 ⇄（蓝）：转换
            Pen ap = Pen(acc, 2.2);
            Line(dc, ap, 18.5, 7.0, 27.5, 7.0);
            Line(dc, ap, 24.7, 4.4, 27.5, 7.0);
            Line(dc, ap, 24.7, 9.6, 27.5, 7.0);
            Line(dc, ap, 27.5, 12.8, 18.5, 12.8);
            Line(dc, ap, 21.3, 10.2, 18.5, 12.8);
            Line(dc, ap, 21.3, 15.4, 18.5, 12.8);
            dc.Pop();
        }

        /// <summary>方案：折角文档（灰描边）+ 文本行 + 蓝色强调行。</summary>
        private static void DrawProfiles(DrawingContext dc, double u, Color fg, Color acc)
        {
            dc.PushTransform(new ScaleTransform(u, u));
            var doc = new StreamGeometry();
            using (StreamGeometryContext sgc = doc.Open())
            {
                sgc.BeginFigure(new Point(10, 4.5), false, true);
                sgc.LineTo(new Point(17.5, 4.5), true, false);
                sgc.LineTo(new Point(23.5, 10.5), true, false);
                sgc.LineTo(new Point(23.5, 26.5), true, false);
                sgc.QuadraticBezierTo(new Point(23.5, 28.5), new Point(21.5, 28.5), true, false);
                sgc.LineTo(new Point(10.5, 28.5), true, false);
                sgc.QuadraticBezierTo(new Point(8.5, 28.5), new Point(8.5, 26.5), true, false);
                sgc.LineTo(new Point(8.5, 6.5), true, false);
                sgc.QuadraticBezierTo(new Point(8.5, 4.5), new Point(10, 4.5), true, false);
            }
            doc.Freeze();
            dc.DrawGeometry(null, Pen(fg, 2.0), doc);
            // 折角线
            var fold = new StreamGeometry();
            using (StreamGeometryContext sgc = fold.Open())
            {
                sgc.BeginFigure(new Point(17.5, 4.5), false, false);
                sgc.LineTo(new Point(17.5, 10.5), true, false);
                sgc.LineTo(new Point(23.5, 10.5), true, false);
            }
            fold.Freeze();
            dc.DrawGeometry(null, Pen(fg, 1.6, 0.9), fold);
            // 文本行 + 蓝色强调行
            Bar(dc, Solid(fg, 0.85), 11.5, 14.5, 9.0, 2.0);
            Bar(dc, Solid(fg, 0.85), 11.5, 18.8, 9.0, 2.0);
            Bar(dc, Solid(acc, 1.0), 11.5, 23.1, 6.5, 2.0);
            dc.Pop();
        }

        /// <summary>存为方案（16px）：折角小文档 + 蓝色加号。</summary>
        private static void DrawSave(DrawingContext dc, double u, Color fg, Color acc)
        {
            dc.PushTransform(new ScaleTransform(u, u));
            var doc = new StreamGeometry();
            using (StreamGeometryContext sgc = doc.Open())
            {
                sgc.BeginFigure(new Point(6, 3), false, true);
                sgc.LineTo(new Point(12, 3), true, false);
                sgc.LineTo(new Point(16, 7), true, false);
                sgc.LineTo(new Point(16, 28), true, false);
                sgc.LineTo(new Point(6, 28), true, false);
            }
            doc.Freeze();
            dc.DrawGeometry(null, Pen(fg, 2.0), doc);
            Pen ap = Pen(acc, 2.4);
            Line(dc, ap, 20, 14.5, 28, 14.5);
            Line(dc, ap, 24, 10.5, 24, 18.5);
            dc.Pop();
        }

        /// <summary>应用/管理方案（16px）：列表行 + 蓝色对勾。</summary>
        private static void DrawManage(DrawingContext dc, double u, Color fg, Color acc)
        {
            dc.PushTransform(new ScaleTransform(u, u));
            Bar(dc, Solid(fg, 0.9), 5, 6, 8, 2.6);
            Bar(dc, Solid(fg, 0.9), 5, 12.5, 8, 2.6);
            Bar(dc, Solid(fg, 0.9), 5, 19, 8, 2.6);
            Bar(dc, Solid(fg, 0.9), 5, 25.5, 6, 2.6);
            var check = new StreamGeometry();
            using (StreamGeometryContext sgc = check.Open())
            {
                sgc.BeginFigure(new Point(18, 14.5), false, false);
                sgc.LineTo(new Point(22, 19), true, false);
                sgc.LineTo(new Point(29, 9.5), true, false);
            }
            check.Freeze();
            dc.DrawGeometry(null, Pen(acc, 2.6), check);
            dc.Pop();
        }

        /// <summary>使用说明：问号（accent 圆环 + fg 笔画）。</summary>
        private static void DrawHelp(DrawingContext dc, double u, Color fg, Color acc)
        {
            dc.PushTransform(new ScaleTransform(u, u));
            dc.DrawEllipse(null, Pen(acc, 2.4), new Point(16, 16), 12, 12);
            var q1 = new StreamGeometry();
            using (StreamGeometryContext sgc = q1.Open())
            {
                // 问号弧钩：从左上弧到中心下折
                sgc.BeginFigure(new Point(12.2, 11.5), false, false);
                sgc.BezierTo(new Point(13.4, 9.2), new Point(18.6, 9.2), new Point(18.6, 12.2), true, false);
                sgc.BezierTo(new Point(18.6, 14.4), new Point(16, 14.9), new Point(16, 17.6), true, false);
            }
            q1.Freeze();
            dc.DrawGeometry(null, Pen(fg, 2.6), q1);
            dc.DrawEllipse(Solid(fg, 0.95), null, new Point(16, 21.8), 1.55, 1.55);
            dc.Pop();
        }

        // ───────── IPictureDisp 转换 ─────────

        [ComImport, Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPictureDisp { }

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern IPictureDisp OleCreatePictureIndirect(ref PICTDESC desc, ref Guid iid, bool ownsPicture);

        [StructLayout(LayoutKind.Sequential)]
        private struct PICTDESC
        {
            public int SizeOfStruct;
            public int PicType;
            public IntPtr HBitmap;
            public IntPtr HPalette;
            public int Unused1, Unused2, Unused3;
        }

        private const int PICTYPE_ICON = 3;
        private static Guid IidPictureDisp = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB");

        /// <summary>WPF 位图 → IPictureDisp。
        /// 关键：走 GetHicon() + PICTYPE_ICON（VSTO 同款配方）——HICON 完整携带 32bpp Alpha，
        /// 功能区按透明渲染；此前 GetHbitmap() 路径的 Alpha 在功能区不可靠，
        /// 透明背景整块变成灰色图层（每个图标一块"灰底板"的根因）。</summary>
        private static object ToPictureDisp(BitmapSource source)
        {
            IntPtr hIcon = IntPtr.Zero;
            try
            {
                hIcon = ToHIcon(source);
                var desc = new PICTDESC
                {
                    SizeOfStruct = Marshal.SizeOf(typeof(PICTDESC)),
                    PicType = PICTYPE_ICON,
                    HBitmap = hIcon,   // 联合体同槽位，此处承载 HICON
                    HPalette = IntPtr.Zero
                };
                // ownsPicture=true：IPicture 接管句柄，成功后无需 DestroyIcon
                return OleCreatePictureIndirect(ref desc, ref IidPictureDisp, true);
            }
            catch
            {
                if (hIcon != IntPtr.Zero) DestroyIcon(hIcon);
                return null;
            }
        }

        [DllImport("user32.dll")] private static extern bool DestroyIcon(IntPtr hIcon);

        private static IntPtr ToHIcon(BitmapSource source)
        {
            int w = source.PixelWidth, h = source.PixelHeight;
            var premultiplied = new byte[w * h * 4];
            source.CopyPixels(premultiplied, w * 4, 0);

            // Pbgra32（预乘）→ 直 Alpha，供 GDI 32bppARGB 使用
            var straight = new byte[w * h * 4];
            for (int i = 0; i < premultiplied.Length; i += 4)
            {
                byte a = premultiplied[i + 3];
                straight[i] = a == 0 ? (byte)0 : (byte)Math.Min(255, premultiplied[i] * 255 / a);
                straight[i + 1] = a == 0 ? (byte)0 : (byte)Math.Min(255, premultiplied[i + 1] * 255 / a);
                straight[i + 2] = a == 0 ? (byte)0 : (byte)Math.Min(255, premultiplied[i + 2] * 255 / a);
                straight[i + 3] = a;
            }

            using (var bmp = new System.Drawing.Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {
                var rect = new System.Drawing.Rectangle(0, 0, w, h);
                var data = bmp.LockBits(rect, System.Drawing.Imaging.ImageLockMode.WriteOnly,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                try { Marshal.Copy(straight, 0, data.Scan0, straight.Length); }
                finally { bmp.UnlockBits(data); }
                return bmp.GetHicon();   // HICON 保留完整 Alpha（区别于 GetHbitmap 丢 Alpha）
            }
        }

        /// <summary>导出 PNG（仅自检/设计核对用）。</summary>
        public static void SavePng(string id, OfficeThemeKind theme, string path)
        {
            RenderTargetBitmap bmp = Render(id, id == "ctxSimilar" || id == "btnProfileSave" || id == "btnProfileManage" ? 16 : 32, theme);
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bmp));
            using (var fs = System.IO.File.Create(path)) encoder.Save(fs);
        }
    }
}
