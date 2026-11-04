using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Microsoft.Win32;

namespace TypeRight.Addin.UI
{
    /// <summary>宿主 Word 当前主题（浅色 / 深灰 / 黑色）。</summary>
    public enum OfficeThemeKind
    {
        Light,
        DarkGray,
        Dark
    }

    /// <summary>
    /// Word 宿主主题检测（设计要求：以 Word 实际使用的主题为准，禁止用 Windows 系统深浅色代替）。
    /// M365 各版本主题存储位置漂移（"UI Theme" 值 / 子键 / 身份漫游键，暗色模式开关另存），
    /// 注册表不可靠 → 以【屏幕实测功能区底色】为真值，注册表仅兜底。
    /// 实测：Word 活动窗口顶部功能区带内稀疏采样，取众数色判亮度；窗口不可见或采样被
    /// 彩色窗口遮挡（饱和度高）时回退注册表。
    /// </summary>
    public static class OfficeTheme
    {
        // ---- 结果缓存（图标与悬浮条共用，避免同一帧内重复抓屏） ----
        private static OfficeThemeKind? _cached;
        private static DateTime _cachedAt = DateTime.MinValue;

        /// <summary>当前宿主主题（5 秒缓存；force=true 强制重测，用于主题刚切换后）。</summary>
        public static OfficeThemeKind Detect(NetOffice.WordApi.Application app, bool force = false)
        {
            if (!force && _cached.HasValue && (DateTime.Now - _cachedAt).TotalSeconds < 5)
                return _cached.Value;

            OfficeThemeKind kind;
            bool sampled;
            OfficeThemeKind byPixels = DetectByPixels(app, out sampled);
            kind = sampled ? byPixels : DetectByRegistry(app);

            _cached = kind;
            _cachedAt = DateTime.Now;
            return kind;
        }

        // ───────── 屏幕实测（真值） ─────────

        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        /// <summary>采样 Word 活动窗口功能区带的底色。ok=false 表示无法可靠采样（窗口隐藏/被彩色窗口遮挡）。</summary>
        private static OfficeThemeKind DetectByPixels(NetOffice.WordApi.Application app, out bool ok)
        {
            ok = false;
            try
            {
                IntPtr hwnd = WordWindow.Get(app);
                if (hwnd == IntPtr.Zero || !IsWindowVisible(hwnd) || IsIconic(hwnd)) return OfficeThemeKind.Light;

                RECT r;
                if (!GetWindowRect(hwnd, out r)) return OfficeThemeKind.Light;
                int w = r.Right - r.Left, h = r.Bottom - r.Top;
                if (w < 300 || h < 300) return OfficeThemeKind.Light;

                // 标题栏/选项卡行永远跟随主题（“永不改页面背景”时暗色主题下文档页面仍是白色，
                // 功能区收起后中带采样会落到页面上误判）→ 自上而下按行采样，取首个可信行。
                using (var bmp = new System.Drawing.Bitmap(1, 1))
                using (var g = System.Drawing.Graphics.FromImage(bmp))
                {
                    foreach (double rowFrac in new[] { 0.03, 0.07, 0.13 })
                    {
                        int y = r.Top + (int)(h * rowFrac);
                        var buckets = new Dictionary<int, int>();
                        for (int fx = 0; fx < 12; fx++)
                        {
                            int x = r.Left + (int)(w * 0.15) + (int)(w * 0.70) * fx / 11;
                            g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(1, 1));
                            System.Drawing.Color c = bmp.GetPixel(0, 0);
                            int key = ((int)c.R >> 4) << 8 | ((int)c.G >> 4) << 4 | ((int)c.B >> 4);
                            int n;
                            buckets[key] = buckets.TryGetValue(key, out n) ? n + 1 : 1;
                        }

                        int bestKey = -1, bestCount = 0, bestR = 0, bestG = 0, bestB = 0;
                        foreach (var kv in buckets)
                        {
                            if (kv.Value > bestCount)
                            {
                                bestCount = kv.Value;
                                bestKey = kv.Key;
                                bestR = ((kv.Key >> 8) & 0xF) * 17;
                                bestG = ((kv.Key >> 4) & 0xF) * 17;
                                bestB = (kv.Key & 0xF) * 17;
                            }
                        }
                        if (bestKey < 0 || bestCount < 7) continue;   // 该行无主导色 → 看下一行

                        double max = Math.Max(bestR, Math.Max(bestG, bestB));
                        double min = Math.Min(bestR, Math.Min(bestG, bestB));
                        double saturation = max <= 0 ? 0 : (max - min) / max;
                        if (saturation > 0.25) continue;               // 被彩色控件/窗口盖住 → 看下一行

                        ok = true;
                        double lum = (bestR * 0.299 + bestG * 0.587 + bestB * 0.114) / 255.0;
                        if (lum < 0.45) return OfficeThemeKind.Dark;      // 黑色主题 ≈ #262626
                        if (lum < 0.75) return OfficeThemeKind.DarkGray;  // 深灰主题 ≈ #595959
                        return OfficeThemeKind.Light;                     // 彩色/白色主题 ≈ #F5F5F5
                    }
                }
                return OfficeThemeKind.Light;
            }
            catch
            {
                return OfficeThemeKind.Light;
            }
        }

        // ───────── 注册表兜底 ─────────

        /// <summary>UI Theme：0=彩色 1=深灰 2=黑色 3=白色 4=跟随系统。不同版本写法不同，两处都查。</summary>
        private static OfficeThemeKind DetectByRegistry(NetOffice.WordApi.Application app)
        {
            try
            {
                string ver = "16.0";
                try { ver = app.Version; } catch { }

                // 写法 A：HKCU\Software\Microsoft\Office\<ver>\Common 下的 "UI Theme" 值（本机实测）
                int? a = ReadInt(Registry.CurrentUser,
                    @"Software\Microsoft\Office\" + ver + @"\Common", "UI Theme");
                // 写法 B：...\Common\UI Theme 子键的默认值
                if (!a.HasValue)
                    a = ReadInt(Registry.CurrentUser,
                        @"Software\Microsoft\Office\" + ver + @"\Common\UI Theme", null);

                if (!a.HasValue) return OfficeThemeKind.Light;
                return MapRegistry(a.Value);
            }
            catch
            {
                return OfficeThemeKind.Light;
            }
        }

        private static int? ReadInt(RegistryKey root, string subKey, string name)
        {
            try
            {
                using (RegistryKey k = root.OpenSubKey(subKey))
                {
                    if (k == null) return null;
                    object v = k.GetValue(name);
                    if (v is int) return (int)v;
                    int parsed;
                    if (v is string && int.TryParse((string)v, out parsed)) return parsed;
                    return null;
                }
            }
            catch { return null; }
        }

        private static OfficeThemeKind MapRegistry(int v)
        {
            switch (v)
            {
                case 1: return OfficeThemeKind.DarkGray;
                case 2: return OfficeThemeKind.Dark;
                case 4: return MapSystemTheme(); // 跟随系统：读 Windows 应用深浅色（两条已知路径）
                default: return OfficeThemeKind.Light; // 0 彩色 / 3 白色 → 浅色系
            }
        }

        private static OfficeThemeKind MapSystemTheme()
        {
            // 常见路径 "Windows NT\...\Personalize" 与本机实测路径 "Windows\...\Personalize" 都查
            foreach (string path in new[]
            {
                @"Software\Microsoft\Windows NT\CurrentVersion\Themes\Personalize",
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            })
            {
                int? light = ReadInt(Registry.CurrentUser, path, "AppsUseLightTheme");
                if (light.HasValue) return light.Value == 0 ? OfficeThemeKind.Dark : OfficeThemeKind.Light;
            }
            return OfficeThemeKind.Light;
        }
    }

    /// <summary>三套主题调色板（悬浮窗全部控件颜色来源；色值参考设计稿深色 + Office 原生浅色）。</summary>
    internal sealed class ThemePalette
    {
        public Color Panel, PanelBorder, Text, TextSecondary;
        public Color InputBg, InputBorder, Hover, Pressed, Accent, AccentHover;
        public Color PopupBg, PopupBorder, ItemHover, Separator, Disabled;

        /// <summary>黑色主题（设计稿）：深灰面板 + 浅色文字。</summary>
        public static ThemePalette Dark()
        {
            return new ThemePalette
            {
                Panel = C(0x2A, 0x2A, 0x2A),
                PanelBorder = C(0x62, 0x62, 0x62),
                Text = C(0xF2, 0xF2, 0xF2),
                TextSecondary = C(0xB8, 0xB8, 0xB8),
                InputBg = C(0x33, 0x33, 0x33),
                InputBorder = C(0x6E, 0x6E, 0x6E),
                Hover = C(0x3D, 0x3D, 0x3D),
                Pressed = C(0x48, 0x48, 0x48),
                Accent = C(0x4D, 0x8E, 0xF7),
                AccentHover = C(0x6B, 0xA3, 0xFF),
                PopupBg = C(0x2F, 0x2F, 0x2F),
                PopupBorder = C(0x62, 0x62, 0x62),
                ItemHover = C(0x3D, 0x3D, 0x3D),
                Separator = C(0x3A, 0x3A, 0x3A),
                Disabled = C(0x80, 0x80, 0x80),
            };
        }

        /// <summary>深灰主题。</summary>
        public static ThemePalette DarkGray()
        {
            return new ThemePalette
            {
                Panel = C(0x4A, 0x4A, 0x4A),
                PanelBorder = C(0x85, 0x85, 0x85),
                Text = C(0xF5, 0xF5, 0xF5),
                TextSecondary = C(0xC9, 0xC9, 0xC9),
                InputBg = C(0x56, 0x56, 0x56),
                InputBorder = C(0x90, 0x90, 0x90),
                Hover = C(0x58, 0x58, 0x58),
                Pressed = C(0x62, 0x62, 0x62),
                Accent = C(0x7A, 0xA5, 0xEC),
                AccentHover = C(0x8F, 0xB6, 0xF3),
                PopupBg = C(0x50, 0x50, 0x50),
                PopupBorder = C(0x85, 0x85, 0x85),
                ItemHover = C(0x5C, 0x5C, 0x5C),
                Separator = C(0x5A, 0x5A, 0x5A),
                Disabled = C(0x9A, 0x9A, 0x9A),
            };
        }

        /// <summary>浅色主题（彩色/白色）：贴近 Word 浅色功能区。</summary>
        public static ThemePalette Light()
        {
            return new ThemePalette
            {
                Panel = C(0xFB, 0xFB, 0xFC),
                PanelBorder = C(0xC9, 0xCD, 0xD4),
                Text = C(0x26, 0x26, 0x26),
                TextSecondary = C(0x6E, 0x71, 0x75),
                InputBg = C(0xFF, 0xFF, 0xFF),
                InputBorder = C(0xC4, 0xC9, 0xD0),
                Hover = C(0xE9, 0xEB, 0xEE),
                Pressed = C(0xDC, 0xDF, 0xE4),
                Accent = C(0x2B, 0x57, 0x9A),
                AccentHover = C(0x2F, 0x63, 0xAC),
                PopupBg = C(0xFF, 0xFF, 0xFF),
                PopupBorder = C(0xC9, 0xCD, 0xD4),
                ItemHover = C(0xED, 0xF1, 0xF7),
                Separator = C(0xE5, 0xE7, 0xEB),
                Disabled = C(0x9A, 0xA0, 0xA6),
            };
        }

        public static ThemePalette Get(OfficeThemeKind kind)
        {
            switch (kind)
            {
                case OfficeThemeKind.Dark: return Dark();
                case OfficeThemeKind.DarkGray: return DarkGray();
                default: return Light();
            }
        }

        private static Color C(int r, int g, int b) { return Color.FromRgb((byte)r, (byte)g, (byte)b); }
    }
}
