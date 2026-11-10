using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace TypeRight.Addin.UI
{
    /// <summary>
    /// 表头下划线零位移填空轻量输入框（Enter 确认，Esc 取消，遵循 DESIGN.md Fluent 令牌）。
    /// </summary>
    internal sealed class UnderlinePromptWindow : Window
    {
        private readonly TextBox _input;
        public string InputText { get; private set; }

        public UnderlinePromptWindow(string labelHint)
        {
            Title = "下划线零位移填空";
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            Topmost = true;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            SizeToContent = SizeToContent.WidthAndHeight;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            FontFamily = new FontFamily("Segoe UI, Microsoft YaHei UI");

            var border = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0xFA, 0xFA, 0xFA)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xD9, 0xD9, 0xD9)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(16, 12, 16, 12),
                Width = 340,
                Margin = new Thickness(12)
            };
            border.Effect = new System.Windows.Media.Effects.DropShadowEffect
            {
                BlurRadius = 14,
                ShadowDepth = 3,
                Opacity = 0.28
            };
            border.MouseLeftButtonDown += (s, e) => { try { DragMove(); } catch { } };

            var stack = new StackPanel();

            string promptTitle = string.IsNullOrEmpty(labelHint)
                ? "输入填入下划线的内容（右侧文字零位移）："
                : "填入「" + labelHint + "」（自动定宽吞空格，右侧零位移）：";

            var lbl = new TextBlock
            {
                Text = promptTitle,
                Foreground = new SolidColorBrush(Color.FromRgb(0x1F, 0x1F, 0x1F)),
                FontSize = 12.5,
                Margin = new Thickness(0, 0, 0, 8)
            };
            stack.Children.Add(lbl);

            _input = new TextBox
            {
                FontSize = 13,
                Padding = new Thickness(6, 4, 6, 4),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0x2B, 0x57, 0x9A)),
                BorderThickness = new Thickness(1.2),
                Background = Brushes.White,
                Foreground = new SolidColorBrush(Color.FromRgb(0x1F, 0x1F, 0x1F))
            };
            _input.KeyDown += OnInputKeyDown;
            stack.Children.Add(_input);

            var btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };

            var btnCancel = new Button
            {
                Content = "取消 (Esc)",
                FontSize = 12,
                Padding = new Thickness(10, 4, 10, 4),
                Margin = new Thickness(0, 0, 8, 0),
                Background = new SolidColorBrush(Color.FromRgb(0xEC, 0xEC, 0xEC)),
                Foreground = new SolidColorBrush(Color.FromRgb(0x1F, 0x1F, 0x1F)),
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            btnCancel.Click += (s, e) => { DialogResult = false; Close(); };

            var btnOk = new Button
            {
                Content = "填入 (Enter)",
                FontSize = 12,
                Padding = new Thickness(12, 4, 12, 4),
                Background = new SolidColorBrush(Color.FromRgb(0x2B, 0x57, 0x9A)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            btnOk.Click += (s, e) => Submit();

            btnRow.Children.Add(btnCancel);
            btnRow.Children.Add(btnOk);
            stack.Children.Add(btnRow);

            border.Child = stack;
            Content = border;

            Loaded += (s, e) => { _input.Focus(); };
        }

        private void OnInputKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                Submit();
            }
            else if (e.Key == Key.Escape)
            {
                e.Handled = true;
                DialogResult = false;
                Close();
            }
        }

        private void Submit()
        {
            if (string.IsNullOrWhiteSpace(_input.Text))
            {
                DialogResult = false;
                Close();
                return;
            }
            InputText = _input.Text.Trim();
            DialogResult = true;
            Close();
        }

        public static string Prompt(string labelHint)
        {
            try
            {
                var win = new UnderlinePromptWindow(labelHint);
                if (win.ShowDialog() == true)
                    return win.InputText;
            }
            catch { }
            return null;
        }
    }
}
