using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using WdApp = NetOffice.WordApi;
using WdApplication = NetOffice.WordApi.Application;
using Paragraph = NetOffice.WordApi.Paragraph;
using NetOffice.WordApi.Enums;
using TypeRight.Addin.Controllers;
using TypeRight.Signature;

namespace TypeRight.Addin.UI
{
    /// <summary>
    /// 浮动格式条（PRD v2 P0-2 / §7.2）：样本段落旁出现；
    /// 取值即时应用（每次一个撤销单元）；宽松勾选实时重算；Esc 或点击别处关闭。
    /// </summary>
    public partial class FloatBar : Window
    {
        private static readonly double[] CommonSizes = { 8, 9, 10.5, 12, 14, 16, 18, 22, 26 };

        private readonly TypeRightController _controller;
        private bool _suppressEvents;
        private Popup _colorPopup;

        public FloatBar(TypeRightController controller)
        {
            InitializeComponent();
            _controller = controller;

            CboSize.Items.Add(new SizeItem(null)); // “—” = 不改
            foreach (double s in CommonSizes) CboSize.Items.Add(new SizeItem(s));
            CboSize.SelectedIndex = 0;

            PreviewKeyDown += OnPreviewKeyDown;
            Deactivated += OnDeactivated;
        }

        public void SetFonts(IList<string> fonts)
        {
            _suppressEvents = true;
            try
            {
                CboFont.Items.Clear();
                foreach (string f in fonts) CboFont.Items.Add(f);
                CboFont.SelectedIndex = -1;
            }
            finally { _suppressEvents = false; }
        }

        public void UpdateMatched(int count, TimeSpan duration)
        {
            TxtCount.Text = count <= 1 ? "仅当前段落" : "已匹配 " + count + " 段";
            TxtPerf.Text = count > 1 ? duration.TotalMilliseconds.ToString("0") + "ms" : "";
        }

        /// <summary>定位到样本段落旁（Word GetPoint 给出屏幕物理像素，需换算 WPF DIP）。</summary>
        public void PositionNear(WdApplication app, Paragraph sample)
        {
            try
            {
                int x, y, w, h;
                WdApp.Window win = app.ActiveWindow;
                win.GetPoint(out x, out y, out w, out h, sample.Range);

                double dpiScale = GetDpiScale();
                double px = x / dpiScale, py = y / dpiScale;
                double ph = h / dpiScale;

                // 优先放段落左侧，不够放右侧；再夹到屏幕内
                double barW = ActualWidth > 0 ? ActualWidth : Width;
                double barH = ActualHeight > 0 ? ActualHeight : 120;
                double left = px - barW - 12;
                if (left < 0) left = px + w / dpiScale + 12;
                double top = py - barH - 6;
                if (top < 0) top = py + ph + 6;

                var screen = System.Windows.Forms.Screen.FromPoint(new System.Drawing.Point(x, y));
                double screenRight = screen.Bounds.Right / dpiScale;
                double screenBottom = screen.Bounds.Bottom / dpiScale;
                if (left + barW > screenRight) left = screenRight - barW - 8;
                if (top + barH > screenBottom) top = Math.Max(8, screenBottom - barH - 8);

                Left = left;
                Top = top;
            }
            catch
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
        }

        private double _dpiScale = 0;

        private double GetDpiScale()
        {
            if (_dpiScale > 0) return _dpiScale;
            try
            {
                var source = System.Windows.Interop.HwndSource.FromVisual(this);
                if (source?.CompositionTarget != null)
                {
                    _dpiScale = source.CompositionTarget.TransformToDevice.M11;
                    return _dpiScale;
                }
            }
            catch { }
            _dpiScale = 1.0;
            return _dpiScale;
        }

        public void ShowBar()
        {
            ResetTransientControls();
            Show();
            Activate();
        }

        public void HideBar()
        {
            try { Hide(); } catch { }
        }

        private void ResetTransientControls()
        {
            _suppressEvents = true;
            try
            {
                CboFont.SelectedIndex = -1;
                CboSize.SelectedIndex = 0;
                BtnBold.IsChecked = false;
                ChkStrict.IsChecked = true;
                ChkIgnoreSize.IsChecked = false;
                ChkIgnoreColor.IsChecked = false;
                ChkIgnoreBold.IsChecked = false;
                ChkIgnoreSize.IsEnabled = false;
                ChkIgnoreColor.IsEnabled = false;
                ChkIgnoreBold.IsEnabled = false;
            }
            finally { _suppressEvents = false; }
        }

        // ---- 事件 ----

        private void OnFontSelected(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressEvents) return;
            string font = CboFont.SelectedItem as string;
            if (string.IsNullOrEmpty(font)) return;
            _controller.ApplyChanges(new FormatChanges { EastAsiaFont = font });
        }

        private void OnSizeSelected(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressEvents) return;
            var item = CboSize.SelectedItem as SizeItem;
            if (item == null || !item.SizePt.HasValue) return; // “—” = 不动
            _controller.ApplyChanges(new FormatChanges { SizePt = item.SizePt });
        }

        private void OnBoldToggled(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            bool? bold = BtnBold.IsChecked;
            if (bold == null || !bold.Value)
            {
                BtnBold.IsChecked = false; // 二态使用
                return;
            }
            _controller.ApplyChanges(new FormatChanges { Bold = true });
        }

        private void OnColorClick(object sender, RoutedEventArgs e)
        {
            var presets = new[]
            {
                new object[] { "黑色", "#000000" }, new object[] { "深蓝", "#1A365D" },
                new object[] { "中蓝", "#2B6CB0" }, new object[] { "深红", "#C0392B" },
                new object[] { "深绿", "#2F855A" }, new object[] { "灰", "#595959" }
            };

            var list = new ListBox { BorderThickness = new Thickness(0), Padding = new Thickness(2) };
            foreach (var p in presets)
            {
                var border = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString((string)p[1])),
                    Width = 18, Height = 18, CornerRadius = new CornerRadius(4),
                    Margin = new Thickness(3), Tag = p[1], Cursor = Cursors.Hand
                };
                border.MouseLeftButtonUp += OnPickColor;
                list.Items.Add(border);
            }

            _colorPopup = new Popup
            {
                Child = new Border
                {
                    Background = Brushes.White, CornerRadius = new CornerRadius(8),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(0xC9, 0xD2, 0xE3)),
                    BorderThickness = new Thickness(1), Padding = new Thickness(6),
                    Child = list
                },
                PlacementTarget = BtnColor, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom,
                StaysOpen = false, AllowsTransparency = true
            };
            _colorPopup.IsOpen = true;
        }

        private void OnPickColor(object sender, MouseButtonEventArgs e)
        {
            var border = (Border)sender;
            _colorPopup.IsOpen = false;
            string hex = (string)border.Tag;
            int rgb = int.Parse(hex.Substring(1), System.Globalization.NumberStyles.HexNumber);
            _controller.ApplyChanges(new FormatChanges { ColorRgb = rgb });
        }

        private void OnStrictChanged(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            bool strict = ChkStrict.IsChecked == true;
            ChkIgnoreSize.IsEnabled = !strict;
            ChkIgnoreColor.IsEnabled = !strict;
            ChkIgnoreBold.IsEnabled = !strict;
            _controller.OnIgnoreChanged(ComputeIgnore());
        }

        private void OnIgnoreToggled(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            _controller.OnIgnoreChanged(ComputeIgnore());
        }

        private FormatSignature.IgnoreDims ComputeIgnore()
        {
            if (ChkStrict.IsChecked == true) return FormatSignature.IgnoreDims.None;
            var d = FormatSignature.IgnoreDims.None;
            if (ChkIgnoreSize.IsChecked == true) d |= FormatSignature.IgnoreDims.SizePt;
            if (ChkIgnoreColor.IsChecked == true) d |= FormatSignature.IgnoreDims.Color;
            if (ChkIgnoreBold.IsChecked == true) d |= FormatSignature.IgnoreDims.Bold;
            return d;
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                _controller.CloseSession();
                e.Handled = true;
            }
        }

        private void OnDeactivated(object sender, EventArgs e)
        {
            // 焦点去了非 Word 的窗口（如其他应用）时收起
            var active = System.Windows.Interop.HwndSource.FromHwnd(GetActiveWindow());
            if (active == null) HideBar();
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        private void OnClose(object sender, RoutedEventArgs e)
        {
            _controller.CloseSession();
        }

        private sealed class SizeItem
        {
            public readonly double? SizePt;

            public SizeItem(double? sizePt) { SizePt = sizePt; }

            public override string ToString()
            {
                return SizePt.HasValue ? SizePt.Value.ToString("0.#") : "—";
            }
        }
    }
}
