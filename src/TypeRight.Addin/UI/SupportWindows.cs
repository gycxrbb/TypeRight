using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using TypeRight.Formatting;

namespace TypeRight.Addin.UI
{
    /// <summary>公式转换失败定位清单（PRD v2 §6.4：M &gt; 0 时提供）。</summary>
    public sealed class FormulaFailuresWindow : Window
    {
        public FormulaFailuresWindow()
        {
            Title = "TypeRight · 公式转换失败清单（原文已保留）";
            Width = 520; Height = 320;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ShowInTaskbar = false;
        }

        public void ShowResults(FormulaStats stats)
        {
            var panel = new DockPanel { Margin = new Thickness(14) };
            var head = new TextBlock
            {
                Text = "以下 " + stats.Failed + " 处未转换（原文无损保留，多为非公式误判，可忽略）：",
                FontSize = 12.5,
                Foreground = new SolidColorBrush(Color.FromRgb(0xB4, 0x53, 0x09)),
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap
            };
            DockPanel.SetDock(head, Dock.Top);
            panel.Children.Add(head);

            var list = new ListBox { FontSize = 12 };
            foreach (string loc in stats.FailedLocations) list.Items.Add(loc);
            panel.Children.Add(list);
            Content = panel;
            Show();
            Activate();
        }
    }

    /// <summary>中文字体列表：常用中文族优先，其后按字母序列出本机字体。</summary>
    public static class FontListProvider
    {
        private static readonly string[] Preferred = { "宋体", "黑体", "楷体", "仿宋", "微软雅黑", "等线", "华文中宋", "华文楷体" };

        public static IList<string> GetFontList()
        {
            var installed = new List<string>();
            try
            {
                using (var fonts = new System.Drawing.Text.InstalledFontCollection())
                {
                    foreach (var f in fonts.Families) installed.Add(f.Name);
                }
            }
            catch { }
            if (installed.Count == 0) return Preferred.ToList();

            var result = new List<string>();
            var set = new HashSet<string>(installed, StringComparer.OrdinalIgnoreCase);
            foreach (string p in Preferred)
            {
                if (set.Contains(p)) result.Add(p);
            }
            foreach (string f in installed.OrderBy(f => f, StringComparer.CurrentCulture))
            {
                if (!result.Contains(f, StringComparer.OrdinalIgnoreCase)) result.Add(f);
            }
            return result;
        }
    }
}
