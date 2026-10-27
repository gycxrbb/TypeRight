using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using WdApp = NetOffice.WordApi;
using WdApplication = NetOffice.WordApi.Application;

namespace TypeRight.Addin.UI
{
    /// <summary>轻提示（PRD v2 §7.3）：Word 窗口右下角，不打断操作。</summary>
    public partial class ToastWindow : Window
    {
        private readonly DispatcherTimer _timer;
        private Action _action;

        public ToastWindow()
        {
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            ShowActivated = false;
            Topmost = true;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            SizeToContent = SizeToContent.WidthAndHeight;
            MouseLeftButtonDown += (s, e) => { };

            _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
            _timer.Tick += (s, e) => { _timer.Stop(); FadeOut(); };
        }

        public void Show(WdApplication app, string message, string actionLabel, Action action, int durationMs)
        {
            _action = action;
            Dispatcher.Invoke(() =>
            {
                Content = BuildContent(message, actionLabel, action != null);
                Position(app);
                Show();
                Opacity = 1;
                _timer.Interval = TimeSpan.FromMilliseconds(Math.Max(1500, durationMs));
                _timer.Start();
            });
        }

        private UIElement BuildContent(string message, string actionLabel, bool hasAction)
        {
            var border = new System.Windows.Controls.Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x22, 0x2A, 0x35)),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(16, 10, 12, 10),
                MaxWidth = 460,
                Margin = new Thickness(10, 6, 12, 12)
            };
            border.Effect = new System.Windows.Media.Effects.DropShadowEffect
            { BlurRadius = 10, ShadowDepth = 2, Opacity = 0.3 };

            var grid = new System.Windows.Controls.Grid();
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });

            var text = new System.Windows.Controls.TextBlock
            {
                Text = message,
                Foreground = Brushes.White,
                FontSize = 12.5,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            };
            System.Windows.Controls.Grid.SetColumn(text, 0);
            grid.Children.Add(text);

            if (hasAction)
            {
                var btn = new System.Windows.Controls.Button
                {
                    Content = actionLabel,
                    Background = new SolidColorBrush(Color.FromRgb(0x2B, 0x5C, 0xB8)),
                    Foreground = Brushes.White,
                    FontSize = 12,
                    Padding = new Thickness(12, 4, 12, 4),
                    BorderThickness = new Thickness(0),
                    Cursor = System.Windows.Input.Cursors.Hand,
                    Margin = new Thickness(12, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center
                };
                btn.Click += (s, e) =>
                {
                    var a = _action;
                    FadeOut();
                    if (a != null) a();
                };
                System.Windows.Controls.Grid.SetColumn(btn, 1);
                grid.Children.Add(btn);
            }

            border.Child = grid;
            return border;
        }

        private void Position(WdApplication app)
        {
            try
            {
                IntPtr hwnd = new IntPtr(app.ActiveWindow.Hwnd);
                RECT r;
                if (GetWindowRect(hwnd, out r))
                {
                    double dpi = GetDpiForWindow(hwnd) / 96.0;
                    Left = r.Right / dpi - Width - 28;
                    Top = r.Bottom / dpi - 160;
                    if (Left < 0) Left = 40;
                    if (Top < 0) Top = 80;
                    return;
                }
            }
            catch { }
            var area = SystemParameters.WorkArea;
            Left = area.Right - 380;
            Top = area.Bottom - 160;
        }

        private void FadeOut()
        {
            Dispatcher.Invoke(() =>
            {
                var anim = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(350));
                anim.Completed += (s, e) => Hide();
                BeginAnimation(OpacityProperty, anim);
            });
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);
    }
}
