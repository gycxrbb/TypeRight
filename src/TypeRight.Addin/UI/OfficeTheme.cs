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
            // 回退按宿主分流：注册表里只有 WORD 的主题键，WPS 读它是跨宿主污染
            //（WPS 亮色 + Word 暗色 → WPS 里错误出深色浮条的根因）
            kind = sampled ? byPixels
                : (WordCompat.IsWps(app) ? OfficeThemeKind.Light : DetectByRegistry(app));

            _cached = kind;
            _cachedAt = DateTime.Now;
            return kind;
        }

        // ───────── 屏幕实测（真值） ─────────

        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
        private const uint PW_RENDERFULLCONTENT = 0x2;   // 分层/DirectComposition 内容也能抓

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

                // 采样改为 PrintWindow 抓宿主窗口自身（2026-11-07 根治）：屏幕 CopyFromScreen 在
                // 宿主被遮挡/非前台时会采到别的窗口（回退链再读 Word 注册表 → WPS 亮色出深色条、
                // Word 被遮挡时注册表键漂移出浅色条）。PrintWindow 与遮挡无关，永远采宿主自己的像素。
                using (var cap = new System.Drawing.Bitmap(w, h))
                using (var gc = System.Drawing.Graphics.FromImage(cap))
                {
                    IntPtr hdc = gc.GetHdc();
                    bool printed = PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT);
                    gc.ReleaseHdc(hdc);
                    if (!printed) return OfficeThemeKind.Light;

                    // 标题栏/选项卡行永远跟随主题（“永不改页面背景”时暗色主题下文档页面仍是白色）
                    foreach (double rowFrac in new[] { 0.03, 0.07, 0.13 })
                    {
                        int y = (int)(h * rowFrac);
                        var buckets = new Dictionary<int, int>();
                        for (int fx = 0; fx < 12; fx++)
                        {
                            int x = (int)(w * 0.15) + (int)(w * 0.70) * fx / 11;
                            System.Drawing.Color c = cap.GetPixel(x, y);
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

        /// <summary>黑色主题：Word 深色 Fluent——中性灰阶、隐边框、唯一蓝强调（2026-11 换皮对齐宿主原生）。</summary>
        public static ThemePalette Dark()
        {
            return new ThemePalette
            {
                Panel = C(0x2B, 0x2B, 0x2F),
                PanelBorder = C(0x3F, 0x3F, 0x44),
                Text = C(0xF3, 0xF3, 0xF3),
                TextSecondary = C(0xAD, 0xAD, 0xB3),
                InputBg = C(0x38, 0x38, 0x3E),
                InputBorder = C(0x57, 0x57, 0x5E),
                Hover = C(0x37, 0x37, 0x3D),
                Pressed = C(0x42, 0x42, 0x4A),
                Accent = C(0x6C, 0xA0, 0xF5),
                AccentHover = C(0x84, 0xB0, 0xF7),
                PopupBg = C(0x33, 0x33, 0x38),
                PopupBorder = C(0x45, 0x45, 0x4B),
                ItemHover = C(0x42, 0x42, 0x4A),
                Separator = C(0x3A, 0x3A, 0x3F),
                Disabled = C(0x77, 0x77, 0x7D),
            };
        }

        /// <summary>深灰主题：Word Dark Gray 档（中间调铬面，偏冷）。</summary>
        public static ThemePalette DarkGray()
        {
            return new ThemePalette
            {
                Panel = C(0x44, 0x44, 0x4C),
                PanelBorder = C(0x5A, 0x5A, 0x63),
                Text = C(0xF5, 0xF5, 0xF5),
                TextSecondary = C(0xC2, 0xC2, 0xC8),
                InputBg = C(0x50, 0x50, 0x5A),
                InputBorder = C(0x6E, 0x6E, 0x78),
                Hover = C(0x4E, 0x4E, 0x57),
                Pressed = C(0x5A, 0x5A, 0x64),
                Accent = C(0x7F, 0xA8, 0xF0),
                AccentHover = C(0x93, 0xBA, 0xF4),
                PopupBg = C(0x4A, 0x4A, 0x52),
                PopupBorder = C(0x62, 0x62, 0x6B),
                ItemHover = C(0x55, 0x55, 0x5F),
                Separator = C(0x55, 0x55, 0x5C),
                Disabled = C(0x9E, 0x9E, 0xA6),
            };
        }

        /// <summary>浅色主题：Word 浅色 Fluent——白面板、发丝边、菜单悬停浅蓝晕（2026-11 换皮）。</summary>
        public static ThemePalette Light()
        {
            return new ThemePalette
            {
                Panel = C(0xFA, 0xFA, 0xFA),
                PanelBorder = C(0xD9, 0xD9, 0xD9),
                Text = C(0x1F, 0x1F, 0x1F),
                TextSecondary = C(0x61, 0x61, 0x61),
                InputBg = C(0xFF, 0xFF, 0xFF),
                InputBorder = C(0xC8, 0xC8, 0xC8),
                Hover = C(0xEC, 0xEC, 0xEC),
                Pressed = C(0xDE, 0xDE, 0xDE),
                Accent = C(0x2B, 0x57, 0x9A),
                AccentHover = C(0x2F, 0x63, 0xAC),
                PopupBg = C(0xFF, 0xFF, 0xFF),
                PopupBorder = C(0xE1, 0xE1, 0xE1),
                ItemHover = C(0xE8, 0xF1, 0xFB),
                Separator = C(0xEA, 0xEA, 0xEA),
                Disabled = C(0xA6, 0xA6, 0xA6),
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
