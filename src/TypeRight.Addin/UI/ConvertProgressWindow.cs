using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using WdApplication = NetOffice.WordApi.Application;

namespace TypeRight.Addin.UI
{
    /// <summary>公式转换进度窗（2026-11 体验优化）：三层反馈——
    /// 即刻（等待光标+状态栏在 Controller）／过程（本窗：阶段文案+计数+进度条+Esc 提示）／
    /// 完成（就地变完成态停 1.8s 淡出，替代右下角 toast）。
    /// 关键工程点：转换在 Word UI 线程上同步运行，动画时钟会冻住——
    /// 过程态只做"阶段性跳变"（改文本/宽度 + UpdateLayout，由 WPF 合成线程上屏），
    /// 淡出动画发生在转换结束后的空闲线程上，不受影响。</summary>
    public sealed class ConvertProgressWindow : Window
    {
        private readonly TextBlock _txtPhase;
        private readonly TextBlock _txtCount;
        private readonly Border _barFill;
        private readonly Border _card;
        private readonly Grid _grid;
        private readonly DispatcherTimer _timer;
        private const double TrackWidth = 252;
        private const double CardWidth = 300;

        public ConvertProgressWindow()
        {
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            ShowActivated = false;
            Topmost = true;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            Width = CardWidth + 24;
            SizeToContent = SizeToContent.Height;
            Focusable = false;
            Title = "TypeRightProgress";   // 自动化取证寻址用（不显示）

            // Fluent 深色卡（与 Toast 同族；通知类浮层在浅色宿主下保持深色是 Office 惯例）
            _card = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x2B, 0x2B, 0x2F)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0x3F, 0x3F, 0x44)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(16, 12, 16, 10),
                Width = CardWidth,
                Margin = new Thickness(6, 5, 6, 9)
            };
            _card.Effect = new System.Windows.Media.Effects.DropShadowEffect
            { BlurRadius = 14, ShadowDepth = 3, Opacity = 0.28 };

            _grid = new Grid();
            _grid.RowDefinitions.Add(new RowDefinition());
            _grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(10) });
            _grid.RowDefinitions.Add(new RowDefinition());
            _grid.RowDefinitions.Add(new RowDefinition());

            _txtPhase = new TextBlock
            {
                Text = "正在准备…", Foreground = new SolidColorBrush(Color.FromRgb(0xF3, 0xF3, 0xF3)),
                FontSize = 13, FontWeight = FontWeights.SemiBold
            };
            Grid.SetRow(_txtPhase, 0);
            _grid.Children.Add(_txtPhase);

            // 进度条（轨道 + 填充；填充宽度直接赋值，不走动画时钟）
            var barGrid = new Grid();
            var track = new Border
            {
                Height = 3, CornerRadius = new CornerRadius(1.5),
                Background = new SolidColorBrush(Color.FromRgb(0x3A, 0x3A, 0x3F)),
                Width = TrackWidth, HorizontalAlignment = HorizontalAlignment.Left
            };
            _barFill = new Border
            {
                Height = 3, CornerRadius = new CornerRadius(1.5),
                Background = new SolidColorBrush(Color.FromRgb(0x6C, 0xA0, 0xF5)),
                Width = 0, HorizontalAlignment = HorizontalAlignment.Left
            };
            barGrid.Children.Add(track);
            barGrid.Children.Add(_barFill);
            Grid.SetRow(barGrid, 1);
            Grid.SetRow(barGrid, 1);
            _grid.Children.Add(barGrid);

            _txtCount = new TextBlock
            {
                Text = "", Foreground = new SolidColorBrush(Color.FromRgb(0xAD, 0xAD, 0xB3)),
                FontSize = 11.5, HorizontalAlignment = HorizontalAlignment.Left
            };
            Grid.SetRow(_txtCount, 2);
            _grid.Children.Add(_txtCount);

            var esc = new TextBlock
            {
                Text = "按 Esc 取消", Foreground = new SolidColorBrush(Color.FromRgb(0xAD, 0xAD, 0xB3)),
                FontSize = 10.5, HorizontalAlignment = HorizontalAlignment.Right
            };
            Grid.SetRow(esc, 3);
            _grid.Children.Add(esc);

            _card.Child = _grid;
            Content = _card;

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(1800) };
            _timer.Tick += (s, e) => { _timer.Stop(); FadeOut(); };
        }

        /// <summary>显示在 Word 当前视口居中（水平正中、垂直约 45%——避开功能区与状态栏，
        /// 转换期间用户视线正落在文档中央，进度反馈在视线内而非角落）。</summary>
        public void ShowAt(WdApplication app)
        {
            try
            {
                IntPtr hwnd = WordWindow.Get(app);
                RECT r;
                if (GetWindowRect(hwnd, out r))
                {
                    double dpi = DpiScale(hwnd);
                    double winW = (r.Right - r.Left) / dpi, winH = (r.Bottom - r.Top) / dpi;
                    Left = r.Left / dpi + (winW - (CardWidth + 24)) / 2;
                    Top = r.Top / dpi + winH * 0.45 - ActualHeight / 2;
                    if (Left < r.Left / dpi + 10) Left = r.Left / dpi + 10;
                    if (Top < r.Top / dpi + 80) Top = r.Top / dpi + 80;
                }
                else PositionFallback();
            }
            catch { PositionFallback(); }
            Show();
            Opacity = 1;
            UpdateLayout();
        }

        private void PositionFallback()
        {
            var area = SystemParameters.WorkArea;
            Left = area.Left + (area.Width - (CardWidth + 24)) / 2;
            Top = area.Top + area.Height * 0.42;
        }

        /// <summary>过程态更新（在 Word UI 线程被转换占用时调用：同步赋值 + 强制布局，
        /// 由 WPF 合成线程上屏——不依赖动画时钟）。total&lt;=0 表示不确定总量（条走脉冲位）。</summary>
        public void Update(string phase, int done, int total)
        {
            _txtPhase.Text = phase ?? "";
            if (total > 0)
            {
                _txtCount.Text = string.Format("{0} / {1}", done, total);
                double frac = Math.Max(0.0, Math.Min(1.0, (double)done / total));
                _barFill.Width = TrackWidth * frac;
            }
            else
            {
                _txtCount.Text = done > 0 ? done.ToString() : "";
                _barFill.Width = TrackWidth * 0.06;   // 不确定态：短条示意进行中
            }
            UpdateLayout();
        }

        /// <summary>完成态：主文案 + 可选副行，停 1.8s 后淡出关闭（此时线程已空闲，动画可用）。</summary>
        public void ShowDone(string main, string sub = null)
        {
            _grid.Children.Clear();
            _txtPhase.Text = main;
            _txtCount.Text = sub ?? "";
            _txtCount.FontSize = 11.5;
            var done = new TextBlock
            {
                Text = _txtPhase.Text, Foreground = _txtPhase.Foreground,
                FontSize = 13, FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(done, 0);
            Grid.SetRowSpan(done, 3);
            _grid.Children.Add(done);
            if (!string.IsNullOrEmpty(sub))
            {
                var subT = new TextBlock
                {
                    Text = sub, Foreground = _txtCount.Foreground, FontSize = 11,
                    TextWrapping = TextWrapping.Wrap
                };
                Grid.SetRow(subT, 3);
                Grid.SetRowSpan(subT, 1);
                _grid.Children.Add(subT);
            }
            _timer.Start();
        }

        /// <summary>异常路径立即关闭。</summary>
        public void CloseNow()
        {
            try { _timer.Stop(); Close(); } catch { }
        }

        private void FadeOut()
        {
            try
            {
                var anim = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(350));
                anim.Completed += (s, e) => { try { Close(); } catch { } };
                BeginAnimation(OpacityProperty, anim);
            }
            catch { try { Close(); } catch { } }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private static double DpiScale(IntPtr hwnd)
        {
            try { return GetDpiForWindow(hwnd) / 96.0; }
            catch
            {
                try { using (var g = System.Drawing.Graphics.FromHwnd(hwnd)) return g.DpiX / 96.0; }
                catch { return 1.0; }
            }
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);
    }
}
