using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
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
    /// 主题跟随宿主 Word（OfficeTheme 屏幕实测），整套控件颜色由主题资源驱动。
    /// </summary>
    public partial class FloatBar : Window
    {
        private static readonly double[] CommonSizes = { 8, 9, 10.5, 12, 14, 16, 18, 22, 26 };

        private readonly TypeRightController _controller;
        private bool _suppressEvents;
        private bool _dragging;    // DragMove 模态移动中（此间 Esc 只取消拖动，不关闭会话）
        private bool _userMoved;   // 用户手动拖过：此后不再回弹到样本段落旁
        private Popup _colorPopup;
        private OfficeThemeKind _themeKind;

        public FloatBar(TypeRightController controller)
        {
            _suppressEvents = true;   // XAML 初始化期间 IsChecked 赋值会触发事件，此时控件尚未全部构建
            InitializeComponent();
            _suppressEvents = false;
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
            // 耗时仅进日志（设计稿要求 UI 不显示调试数据）
            TxtCount.Text = count <= 1 ? "仅当前段落" : "已匹配 " + count + " 段";
        }

        /// <summary>定位到样本段落旁（Word GetPoint 给出屏幕物理像素，需换算 WPF DIP）。
        /// 用户手动拖动过则保持其位置，仅夹回可见屏幕。</summary>
        public void PositionNear(WdApplication app, Paragraph sample)
        {
            if (_userMoved) return;
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
                double barH = ActualHeight > 0 ? ActualHeight : 140;
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

        // ───────── 主题（跟随宿主 Word） ─────────

        /// <summary>按主题覆盖全部控件颜色（仅 UI 资源，不触碰文档与控件状态）。</summary>
        public void ApplyTheme(OfficeThemeKind kind)
        {
            _themeKind = kind;
            ThemePalette p = ThemePalette.Get(kind);
            SetBrush("PanelBrush", p.Panel);
            SetBrush("PanelBorderBrush", p.PanelBorder);
            SetBrush("TextBrush", p.Text);
            SetBrush("TextSecondaryBrush", p.TextSecondary);
            SetBrush("InputBgBrush", p.InputBg);
            SetBrush("InputBorderBrush", p.InputBorder);
            SetBrush("HoverBrush", p.Hover);
            SetBrush("PressedBrush", p.Pressed);
            SetBrush("AccentBrush", p.Accent);
            SetBrush("AccentHoverBrush", p.AccentHover);
            SetBrush("PopupBgBrush", p.PopupBg);
            SetBrush("PopupBorderBrush", p.PopupBorder);
            SetBrush("ItemHoverBrush", p.ItemHover);
            SetBrush("SeparatorBrush", p.Separator);
            SetBrush("DisabledBrush", p.Disabled);
        }

        private void SetBrush(string key, Color c)
        {
            var b = new SolidColorBrush(c);
            b.Freeze();
            Resources[key] = b;
        }

        private Brush BrushOf(string key)
        {
            return TryFindResource(key) as Brush ?? Brushes.White;
        }

        /// <summary>底部提示与真实配置一致：联动开 → “西文自动联动 X”；关 → 如实说明。</summary>
        private void RefreshHints()
        {
            try
            {
                var s = _controller.Settings;
                TxtHint.Text = s.LatinLinkageEnabled
                    ? "西文自动联动 " + s.DefaultLatinFont
                    : "西文保持原样（联动已关）";
            }
            catch { }
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
            try { ApplyTheme(OfficeTheme.Detect(_controller.App)); }
            catch { ApplyTheme(OfficeThemeKind.Light); }
            RefreshHints();
            if (_userMoved) EnsureOnScreen();  // 分辨率/显示器变化后防止停在屏外
            ResetTransientControls();
            TryAttachToWordWindow();
            Show();
            Activate();
        }

        public void HideBar()
        {
            try { Hide(); } catch { }
        }

        /// <summary>归属 Word 活动窗口：随 Word 最小化/置底，不作为全局置顶遮挡其他程序。</summary>
        private void TryAttachToWordWindow()
        {
            try
            {
                IntPtr hwnd = new IntPtr(_controller.App.ActiveWindow.Hwnd);
                if (hwnd != IntPtr.Zero)
                {
                    var helper = new WindowInteropHelper(this);
                    if (helper.Owner != hwnd) helper.Owner = hwnd;
                }
            }
            catch { }
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
                SetColorStrip(null);
            }
            finally { _suppressEvents = false; }
        }

        private void SetColorStrip(int? rgb)
        {
            var strip = FindName("ColorStrip") as Border;
            if (strip == null) return;
            var c = rgb.HasValue ? Color.FromRgb((byte)(rgb.Value >> 16), (byte)(rgb.Value >> 8), (byte)rgb.Value)
                                 : Color.FromRgb(0, 0, 0);
            var b = new SolidColorBrush(c);
            b.Freeze();
            strip.Background = b;
        }

        // ---- 事件 ----

        /// <summary>整个条身任意非控件位置按住左键即可拖动（标题行为显式把手）。
        /// 按钮/下拉/复选框自行处理鼠标按下（e.Handled），不会误触拖动。</summary>
        private void OnBarDragMove(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left || _dragging) return;

            double beforeLeft = Left, beforeTop = Top;
            _dragging = true;
            try { DragMove(); }
            catch { } // 按下后立即松开等边界情况，DragMove 会抛 InvalidOperationException
            finally { _dragging = false; }

            // Esc 取消拖动时 DragMove 会弹回原位；只有真实移动过才视为用户接管了位置
            if (Math.Abs(Left - beforeLeft) > 0.5 || Math.Abs(Top - beforeTop) > 0.5) _userMoved = true;
        }

        /// <summary>用户拖动后的位置在下次显示时夹回可见屏幕（多显示器/改分辨率后防跑出屏外）。</summary>
        private void EnsureOnScreen()
        {
            try
            {
                double scale = GetDpiScale();
                var screen = System.Windows.Forms.Screen.FromPoint(
                    new System.Drawing.Point((int)(Left * scale), (int)(Top * scale)));
                double monLeft = screen.Bounds.Left / scale, monTop = screen.Bounds.Top / scale;
                double monRight = screen.Bounds.Right / scale, monBottom = screen.Bounds.Bottom / scale;
                double w = ActualWidth > 0 ? ActualWidth : Width;
                double h = ActualHeight > 0 ? ActualHeight : 140;
                if (Left + w > monRight) Left = Math.Max(monLeft, monRight - w - 8);
                if (Top + h > monBottom) Top = Math.Max(monTop, monBottom - h - 8);
                if (Left < monLeft) Left = monLeft + 8;
                if (Top < monTop) Top = monTop + 8;
            }
            catch { }
        }

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

            // 调色板（主题化：面板底色/描边/悬停环）
            Brush popupBg = BrushOf("PopupBgBrush"), popupBorder = BrushOf("PopupBorderBrush"),
                 accent = BrushOf("AccentBrush"), panelBorder = BrushOf("PanelBorderBrush");

            var row = new StackPanel { Orientation = Orientation.Horizontal };
            foreach (var p in presets)
            {
                var swatch = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString((string)p[1])),
                    Width = 20, Height = 20, CornerRadius = new CornerRadius(3),
                    Margin = new Thickness(3), BorderThickness = new Thickness(1),
                    BorderBrush = panelBorder, Cursor = Cursors.Hand, ToolTip = (string)p[0]
                };
                swatch.MouseEnter += (s, _) => { swatch.BorderBrush = accent; swatch.BorderThickness = new Thickness(2); };
                swatch.MouseLeave += (s, _) => { swatch.BorderBrush = panelBorder; swatch.BorderThickness = new Thickness(1); };
                swatch.MouseLeftButtonUp += OnPickColor;
                row.Children.Add(swatch);
            }

            _colorPopup = new Popup
            {
                Child = new Border
                {
                    Background = popupBg, CornerRadius = new CornerRadius(3),
                    BorderBrush = popupBorder,
                    BorderThickness = new Thickness(1), Padding = new Thickness(5),
                    Child = row
                },
                PlacementTarget = BtnColor, Placement = PlacementMode.Bottom,
                StaysOpen = false, AllowsTransparency = true
            };
            _colorPopup.IsOpen = true;
        }

        private void OnPickColor(object sender, MouseButtonEventArgs e)
        {
            var border = (Border)sender;
            if (_colorPopup != null) _colorPopup.IsOpen = false;
            string hex = (string)border.Tag ?? ColorToHex(border.Background as SolidColorBrush);
            int rgb = int.Parse(hex.Substring(1), System.Globalization.NumberStyles.HexNumber);
            SetColorStrip(rgb);
            _controller.ApplyChanges(new FormatChanges { ColorRgb = rgb });
        }

        private static string ColorToHex(SolidColorBrush b)
        {
            if (b == null) return "#000000";
            return "#" + b.Color.R.ToString("X2") + b.Color.G.ToString("X2") + b.Color.B.ToString("X2");
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
                if (_dragging) { e.Handled = true; return; } // 拖动中 Esc 仅取消本次拖动（DragMove 弹回原位）
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
