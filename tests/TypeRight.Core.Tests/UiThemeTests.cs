using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Xunit;

namespace TypeRight.Tests
{
    /// <summary>
    /// 主题一致性回归（2026-11 事故：宿主深色主题下，窗口背景写死白色 + 文字取深色令牌
    /// → 标题浅字落白底，"消失"看不见）。
    ///
    /// 这组用例把约束固化下来：三套主题的每一组前景/背景都必须满足 WCAG 对比度，
    /// 且深色主题下卡片里的每个 TextBlock 都不得出现低对比文字。
    ///
    /// 版本：xUnit 需要 STA 线程才能建 WPF 控件；这里用纯色值断言 + 模板着色断言，
    /// 不启动窗口，避免测试宿主受 STA/桌面会话限制。
    /// </summary>
    /// <summary>轻量 RGB：本文件只做色值算术，无需引入整个 WPF。</summary>
    public struct Rgb
    {
        public byte R, G, B;
        public Rgb(byte r, byte g, byte b) { R = r; G = g; B = b; }
    }

    public class UiThemeTests
    {
        // 与 OfficeTheme.cs 的 ThemePalette 同源：直接从源码解析，杜绝手抄漂移
        private static Dictionary<string, Rgb> ParsePalette(string method)
        {
            string path = FindOfficeTheme();
            string src = File.ReadAllText(path);
            int start = src.IndexOf("public static ThemePalette " + method + "()");
            Assert.True(start >= 0, "找不到主题方法: " + method);
            int end = src.IndexOf("return new ThemePalette", start);
            int close = src.IndexOf("};", end);
            string body = src.Substring(end, close - end);

            var map = new Dictionary<string, Rgb>();
            foreach (Match m in Regex.Matches(body, @"(\w+)\s*=\s*C\(0x([0-9A-Fa-f]{2}),\s*0x([0-9A-Fa-f]{2}),\s*0x([0-9A-Fa-f]{2})\)"))
            {
                map[m.Groups[1].Value] = new Rgb(
                    Convert.ToByte(m.Groups[2].Value, 16),
                    Convert.ToByte(m.Groups[3].Value, 16),
                    Convert.ToByte(m.Groups[4].Value, 16));
            }
            Assert.True(map.Count >= 15, method + " 调色板解析不完整，只拿到 " + map.Count + " 项");
            return map;
        }

        private static string FindOfficeTheme()
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null)
            {
                string p = Path.Combine(dir.FullName, "src", "TypeRight.Addin", "UI", "OfficeTheme.cs");
                if (File.Exists(p)) return p;
                dir = dir.Parent;
            }
            Assert.True(false, "找不到 OfficeTheme.cs");
            return null;
        }

        private static double Lum(Rgb c)
        {
            Func<double, double> f = v => { v /= 255.0; return v <= 0.03928 ? v / 12.92 : Math.Pow((v + 0.055) / 1.055, 2.4); };
            return 0.2126 * f(c.R) + 0.7152 * f(c.G) + 0.0722 * f(c.B);
        }

        private static double Contrast(Rgb a, Rgb b)
        {
            double la = Lum(a), lb = Lum(b);
            return (Math.Max(la, lb) + 0.05) / (Math.Min(la, lb) + 0.05);
        }

        public static IEnumerable<object[]> Themes()
        {
            yield return new object[] { "Light" };
            yield return new object[] { "Dark" };
            yield return new object[] { "DarkGray" };
        }

        [Theory]
        [MemberData(nameof(Themes))]
        public void PrimaryTextHasEnoughContrast(string theme)
        {
            var p = ParsePalette(theme);
            // 标题/正文写在窗口底色上
            Assert.True(Contrast(p["Text"], p["Panel"]) >= 4.5,
                theme + " 主文字与窗口底对比度不足：" + Contrast(p["Text"], p["Panel"]).ToString("0.00"));
            // 卡片内文字写在卡片底色（InputBg）上
            Assert.True(Contrast(p["Text"], p["InputBg"]) >= 4.5,
                theme + " 主文字与卡片底对比度不足：" + Contrast(p["Text"], p["InputBg"]).ToString("0.00"));
        }

        [Theory]
        [MemberData(nameof(Themes))]
        public void SecondaryAndAccentAreVisible(string theme)
        {
            var p = ParsePalette(theme);
            Assert.True(Contrast(p["TextSecondary"], p["InputBg"]) >= 3.0,
                theme + " 次级文字对比度不足：" + Contrast(p["TextSecondary"], p["InputBg"]).ToString("0.00"));
            Assert.True(Contrast(p["Accent"], p["InputBg"]) >= 3.0,
                theme + " 强调色对比度不足：" + Contrast(p["Accent"], p["InputBg"]).ToString("0.00"));
        }

        /// <summary>状态色必须随主题变化——深色底上用深红几乎看不见。</summary>
        [Theory]
        [MemberData(nameof(Themes))]
        public void StatusColorsSurviveBothThemes(string theme)
        {
            var p = ParsePalette(theme);
            Assert.True(p.ContainsKey("Success"), theme + " 缺少 Success 状态色");
            Assert.True(p.ContainsKey("Danger"), theme + " 缺少 Danger 状态色");
            Assert.True(Contrast(p["Success"], p["Panel"]) >= 3.0,
                theme + " 成功色对比度不足：" + Contrast(p["Success"], p["Panel"]).ToString("0.00"));
            Assert.True(Contrast(p["Danger"], p["Panel"]) >= 3.0,
                theme + " 错误色对比度不足：" + Contrast(p["Danger"], p["Panel"]).ToString("0.00"));
        }

        /// <summary>深色主题的窗口底与文字必须是"浅字深底"，不能是白底。</summary>
        [Fact]
        public void DarkThemeIsActuallyDark()
        {
            var dark = ParsePalette("Dark");
            Assert.True(Lum(dark["Panel"]) < 0.2, "深色主题的窗口底不够深");
            Assert.True(Lum(dark["Text"]) > 0.8, "深色主题的主文字不够亮");
        }

        // ───────────────── 源码级约束（防回归）─────────────────
        //
        // 2026-11 连续踩了三次同一类坑，全部发生在"写了主题代码但没接线"的地方：
        //   ① 窗口背景写死 Brushes.White
        //   ② 子控件用 new SolidColorBrush(P.Xxx) 静态快照，ApplyTheme 换令牌也不跟随
        //   ③ ApplyTheme 里 `Background = new ...` 直接赋值，覆盖掉构造函数里的
        //      SetResourceReference 绑定 → 之后换主题永远不生效
        //
        // 这三条都靠"看代码"容易漏，用源码断言钉死。

        private static string ReadUi(string file)
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null)
            {
                string p = Path.Combine(dir.FullName, "src", "TypeRight.Addin", "UI", file);
                if (File.Exists(p)) return File.ReadAllText(p);
                dir = dir.Parent;
            }
            Assert.True(false, "找不到 UI 文件: " + file);
            return null;
        }

        [Fact]
        public void ThemedWindow_BindsBackgroundAsResource()
        {
            string src = ReadUi("ThemedWindow.cs");
            // 窗口底必须通过资源引用绑定，而不是留空（留空 = WPF 默认纯白 = 事故 ①）
            Assert.Contains("UseResource(this, Control.BackgroundProperty", src);
        }

        [Fact]
        public void ThemedWindow_ApplyThemeDoesNotClobberResourceBinding()
        {
            string src = ReadUi("ThemedWindow.cs");
            int applyAt = src.IndexOf("public void ApplyTheme(");
            Assert.True(applyAt > 0, "找不到 ApplyTheme");
            int end = src.IndexOf("OnThemeApplied();", applyAt);
            Assert.True(end > applyAt, "ApplyTheme 结构异常");

            string body = src.Substring(applyAt, end - applyAt);
            // 只查代码行，跳过注释（注释里可能提到这个反例本身）
            foreach (string rawLine in body.Split('\n'))
            {
                string line = rawLine.Trim();
                if (line.StartsWith("//") || line.StartsWith("///") || line.StartsWith("*")) continue;
                // ApplyTheme 内部不能直接给 Background 赋值——那会顶掉 SetResourceReference（事故 ③）
                Assert.False(line.Contains("Background = new SolidColorBrush"),
                    "ApplyTheme 内直接给 Background 赋值会覆盖资源绑定：" + line);
                Assert.False(line.Contains("TextElement.SetForeground("),
                    "ApplyTheme 内直接设文字色会脱离主题：" + line);
            }
        }

        [Fact]
        public void ThemedWindow_RefreshesResourcesByRemoveThenAdd()
        {
            string src = ReadUi("ThemedWindow.cs");
            // 可见后同键直接赋值不会让 DynamicResource 重新求值，必须 Remove 再 Add
            Assert.Contains("Resources.Remove(key)", src);
            Assert.Contains("Resources.Add(key, b)", src);
        }

        [Theory]
        [InlineData("SaveProfileWindow.cs")]
        [InlineData("ProfileManagerWindow.cs")]
        [InlineData("ThemedWindow.cs")]
        public void Windows_HaveNoHardcodedWhiteBackground(string file)
        {
            string src = ReadUi(file);
            Assert.DoesNotContain("Background = Brushes.White", src);
        }

        [Theory]
        [InlineData("SaveProfileWindow.cs")]
        [InlineData("ProfileManagerWindow.cs")]
        public void Windows_DoNotUseStaticPaletteSnapshots(string file)
        {
            string src = ReadUi(file);
            // 静态快照不会跟随主题切换（事故 ②）
            Assert.DoesNotContain("new SolidColorBrush(P.", src);
            Assert.DoesNotContain("Color.FromRgb(", src);
        }

        /// <summary>两个窗口都必须继承 ThemedWindow，否则拿不到主题能力。</summary>
        [Theory]
        [InlineData("SaveProfileWindow.cs", "SaveProfileWindow")]
        [InlineData("ProfileManagerWindow.cs", "ProfileManagerWindow")]
        [InlineData("ProfileManagerWindow.cs", "InputDialog")]
        public void Windows_InheritThemedWindow(string file, string className)
        {
            string src = ReadUi(file);
            string pattern = "class " + className + " : ThemedWindow";
            Assert.True(src.Contains(pattern) || src.Contains("class " + className + " : ThemedWindow"),
                className + " 未继承 ThemedWindow");
        }
    }
}
