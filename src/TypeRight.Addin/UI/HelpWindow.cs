using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using TypeRight.Addin.Controllers;

namespace TypeRight.Addin.UI
{
    /// <summary>使用说明窗（发布支持）：三步用法 + 诊断日志入口 + 版本信息。
    /// 远程排障依赖日志——出问题时让用户点"复制诊断信息"发过来即可。</summary>
    public sealed class HelpWindow : Window
    {
        private readonly TypeRightController _controller;

        public HelpWindow(TypeRightController controller)
        {
            _controller = controller;
            Title = "TypeRight 使用说明";
            WindowStyle = WindowStyle.ToolWindow;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.WidthAndHeight;
            MaxWidth = 520;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ShowInTaskbar = false;
            Background = new SolidColorBrush(Color.FromRgb(0xF7, 0xF8, 0xFA));

            Content = BuildContent();
        }

        private string VersionText()
        {
            try
            {
                return typeof(HelpWindow).Assembly.GetName().Version.ToString(3);
            }
            catch { return "?"; }
        }

        private UIElement BuildContent()
        {
            var panel = new StackPanel { Margin = new Thickness(22, 16, 22, 16) };

            panel.Children.Add(new TextBlock
            {
                Text = "TypeRight 指认式批量排版  v" + VersionText(),
                FontSize = 15, FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(0x1F, 0x29, 0x37)),
                Margin = new Thickness(0, 0, 0, 10)
            });

            panel.Children.Add(MakeStep("①", "光标点进一个段落（或选中一段文字）"));
            panel.Children.Add(MakeStep("②", "功能区 TypeRight 选项卡 →「同类选择」"));
            panel.Children.Add(MakeStep("③", "在弹出的浮动条上改字体/字号/加粗…（作用于全部同类段落）"));
            panel.Children.Add(new TextBlock
            {
                Text = "小提示：Esc 或点击其他位置关闭浮动条；「下一处 ⤵」把命中处设为真选区，可用 Word 原生功能调格式；浮条上「匹配」行可切换宽松模式（忽略字号/颜色/加粗）。",
                FontSize = 11.5, Foreground = new SolidColorBrush(Color.FromRgb(0x5B, 0x66, 0x73)),
                TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 10, 0, 4), LineHeight = 18
            });
            panel.Children.Add(new TextBlock
            {
                Text = "公式转换：一键把 ax²+bx+c=0 这类纯文本公式全文转为可编辑公式对象（TC 级失败片段原样保留）。",
                FontSize = 11.5, Foreground = new SolidColorBrush(Color.FromRgb(0x5B, 0x66, 0x73)),
                TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 12), LineHeight = 18
            });

            var sep = new Separator { Margin = new Thickness(0, 0, 0, 10) };
            panel.Children.Add(sep);

            var btns = new StackPanel { Orientation = Orientation.Horizontal };
            var btnLog = MakeButton("打开日志文件夹", 128, () =>
            {
                try { TypeRightAddin.OpenLogFolder(); } catch { }
            });
            var btnCopy = MakeButton("复制诊断信息", 128, () =>
            {
                try
                {
                    System.Windows.Forms.Clipboard.SetText(TypeRightAddin.CollectDiagnostics(_controller.App));
                    MessageBox.Show(this, "诊断信息已复制（版本 + 最近日志），直接粘贴发给作者即可。", "TypeRight",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "复制失败：" + ex.Message, "TypeRight", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            });
            var btnClose = MakeButton("关闭", 70, Close);
            btns.Children.Add(btnLog);
            btns.Children.Add(btnCopy);
            btns.Children.Add(btnClose);
            panel.Children.Add(btns);

            var border = new Border
            {
                Background = Brushes.White,
                CornerRadius = new CornerRadius(6),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xD4, 0xDA, 0xE0)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(2),
                Child = panel
            };
            return border;
        }

        private static UIElement MakeStep(string no, string text)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 3, 0, 3) };
            sp.Children.Add(new TextBlock
            {
                Text = no, FontSize = 13, FontWeight = FontWeights.Bold, Width = 26,
                Foreground = new SolidColorBrush(Color.FromRgb(0x2B, 0x6C, 0xB0)),
                VerticalAlignment = VerticalAlignment.Top
            });
            sp.Children.Add(new TextBlock
            {
                Text = text, FontSize = 12.5, TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush(Color.FromRgb(0x1F, 0x29, 0x37)), LineHeight = 19
            });
            return sp;
        }

        private static Button MakeButton(string text, double width, Action onClick)
        {
            var b = new Button
            {
                Content = text, FontSize = 12, Padding = new Thickness(10, 5, 10, 5),
                Margin = new Thickness(0, 0, 8, 0), Width = width
            };
            b.Click += (s, e) => onClick();
            return b;
        }
    }
}
