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
        /// <summary>Word"开始"选项卡字号列表（中西文常用字号全集；A± 沿此表步进）。</summary>
        private static readonly double[] SizeSteps = { 5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 15, 16, 18, 20, 22, 24, 26, 28, 30, 32, 36, 42, 48, 72 };

        /// <summary>当前字号（样本锚点字号；每次应用后更新——"字号提示"）。</summary>
        private double? _currentSizePt;

        private readonly TypeRightController _controller;
        private bool _suppressEvents;
        private bool _dragging;    // DragMove 模态移动中（此间 Esc 只取消拖动，不关闭会话）
        private bool _userMoved;   // 用户手动拖过：此后不再回弹到样本段落旁
        private Popup _colorPopup;
        private Popup _sizePopup;
        private OfficeThemeKind _themeKind;

        public FloatBar(TypeRightController controller)
        {
            _suppressEvents = true;   // XAML 初始化期间 IsChecked 赋值会触发事件，此时控件尚未全部构建
            InitializeComponent();
            _suppressEvents = false;
            _controller = controller;
            UpdateEmphasisSwatch();   // 色块初值 = 用户设置

            PreviewKeyDown += OnPreviewKeyDown;
            Deactivated += OnDeactivated;
            Closed += (s, e) => EscHook.Uninstall();
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

        public void UpdateMatched(int count, TimeSpan duration, bool isLead = false)
        {
            // 耗时仅进日志（设计稿要求 UI 不显示调试数据）
            if (isLead)
                TxtCount.Text = count <= 1 ? "仅当前引导句" : "已匹配 " + count + " 处引导句";
            else
                TxtCount.Text = count <= 1 ? "仅当前段落" : "已匹配 " + count + " 处";
        }

        /// <summary>定位：优先出现在选中文本/样本段首个可见字符的右下方（鼠标附近）；
        /// 上下左右溢出时向 Word 窗口内侧收；取不到坐标时 Word 窗口居中。
        /// GetPoint 返回屏幕物理像素，需换算 WPF DIP。</summary>
        public void PositionNear(WdApplication app, Paragraph sample)
        {
            if (_userMoved) return;
            try
            {
                double scale = GetDpiScale();

                // 锚点 = 样本段首个可见字符（比整段 GetPoint 稳：长段的矩形会给出奇怪坐标）
                int anchorStart = sample.Range.Start;
                string text = sample.Range.Text ?? "";
                int limit = Math.Min(text.Length, 16);
                for (int i = 0; i < limit; i++)
                {
                    char ch = text[i];
                    if (ch == '\r' || ch == '\a' || ch == '\v' || ch == ' ' || ch == '　' || ch == '\t') continue;
                    anchorStart = sample.Range.Start + i;
                    break;
                }

                int x, y, w, h;
                WdApp.Window win = app.ActiveWindow;
                win.GetPoint(out x, out y, out w, out h, sample.Range.Document.Range(anchorStart, anchorStart + 1));

                double px = x / scale, py = y / scale, ph = h / scale;
                double barW = ActualWidth > 0 ? ActualWidth : Width;
                double barH = ActualHeight > 0 ? ActualHeight : 150;

                // 锚点右下方；右溢出向左平移，下溢出改放上方；最后夹进屏幕
                double left = px + 18;
                double top = py + ph + 8;
                if (top + barH > ScreenBottom(scale) - 8) top = py - barH - 8;

                var screen = System.Windows.Forms.Screen.FromPoint(new System.Drawing.Point(x, y));
                double monLeft = screen.Bounds.Left / scale, monRight = screen.Bounds.Right / scale;
                if (left + barW > monRight) left = monRight - barW - 12;
                if (left < monLeft) left = monLeft + 12;
                if (top < 0) top = 8;

                Left = left;
                Top = top;
            }
            catch
            {
                try { CenterOverWordWindow(app); }
                catch { WindowStartupLocation = WindowStartupLocation.CenterScreen; }
            }
        }

        private double ScreenBottom(double scale)
        {
            var s = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position);
            return s.Bounds.Bottom / scale;
        }

        /// <summary>兜底：Word 活动窗口中央偏下（不是屏幕中央——多显示器下更贴近视线）。</summary>
        private void CenterOverWordWindow(WdApplication app)
        {
            IntPtr hwnd = WordWindow.Get(app);
            if (hwnd == IntPtr.Zero) { WindowStartupLocation = WindowStartupLocation.CenterScreen; return; }
            var r = new Win32Rect();
            if (!GetWindowRect(hwnd, out r)) { WindowStartupLocation = WindowStartupLocation.CenterScreen; return; }
            double scale = GetDpiScale();
            double monW = System.Windows.Forms.Screen.FromPoint(
                new System.Drawing.Point((r.Left + r.Right) / 2, (r.Top + r.Bottom) / 2)).Bounds.Width / scale;
            double barW = ActualWidth > 0 ? ActualWidth : Width;
            double barH = ActualHeight > 0 ? ActualHeight : 150;
            // 窗口矩形是物理像素；宽度取显示器逻辑宽以容多显示器负坐标
            double wordCenterX = (r.Left + r.Right) / 2.0 / scale;
            Left = wordCenterX - barW / 2;
            if (Left + barW > monW) Left = monW - barW - 12;
            Top = (r.Top + r.Bottom) / 2.0 / scale - barH / 2 + 80;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out Win32Rect rect);

        [StructLayout(LayoutKind.Sequential)]
        private struct Win32Rect { public int Left, Top, Right, Bottom; }

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
            EscHook.Install(this);
        }

        public void HideBar()
        {
            EscHook.Uninstall();
            try { Hide(); } catch { }
        }

        /// <summary>Esc 关闭（提示"Esc 关闭"必须真的好用）：浮条可见期间挂低级键盘钩子，
        /// Esc 且前台窗口属于本 Word 进程时消费并结束会话；焦点在浮条内时走原有 PreviewKeyDown。</summary>
        private static class EscHook
        {
            private const int WH_KEYBOARD_LL = 13;
            private const uint WM_KEYDOWN = 0x0100, WM_SYSKEYDOWN = 0x0104;
            private const int VK_ESCAPE = 0x1B;

            private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);
            private static HookProc _proc;
            private static IntPtr _hook;
            private static WeakReference<FloatBar> _bar;

            [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint threadId);
            [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hhk);
            [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
            [DllImport("kernel32.dll")] private static extern IntPtr GetModuleHandle(IntPtr moduleName);
            [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
            [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

            [StructLayout(LayoutKind.Sequential)]
            private struct KbdLlHookStruct { public uint VkCode, ScanCode, Flags, Time; public IntPtr ExtraInfo; }

            public static void Install(FloatBar bar)
            {
                _bar = new WeakReference<FloatBar>(bar);
                if (_hook != IntPtr.Zero) return;
                try
                {
                    _proc = OnHook;
                    _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(IntPtr.Zero), 0);
                }
                catch { _hook = IntPtr.Zero; }
            }

            public static void Uninstall()
            {
                if (_hook == IntPtr.Zero) return;
                try { UnhookWindowsHookEx(_hook); } catch { }
                _hook = IntPtr.Zero;
            }

            private static IntPtr OnHook(int code, IntPtr wParam, IntPtr lParam)
            {
                if (code == 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
                {
                    var kb = (KbdLlHookStruct)Marshal.PtrToStructure(lParam, typeof(KbdLlHookStruct));
                    if (kb.VkCode == VK_ESCAPE)
                    {
                        FloatBar bar = null;
                        bool alive = _bar != null && _bar.TryGetTarget(out bar);
                        if (alive && bar.IsVisible && !bar.IsDragging)
                        {
                            uint pid;
                            GetWindowThreadProcessId(GetForegroundWindow(), out pid);
                            if (pid == CurrentProcessId)
                            {
                                // 关闭并消费按键（返回 1 = Word 不再收到这枚 Esc）
                                try { bar.Dispatcher.BeginInvoke(new Action(() => bar.EscapeClose())); } catch { }
                                return (IntPtr)1;
                            }
                        }
                    }
                }
                return CallNextHookEx(_hook, code, wParam, lParam);
            }

            private static uint CurrentProcessId
            {
                get { return (uint)System.Diagnostics.Process.GetCurrentProcess().Id; }
            }
        }

        /// <summary>Esc 钩子入口（拖动中留给 DragMove 自己取消，不重复处理）。</summary>
        private void EscapeClose()
        {
            if (_dragging) return;
            try { _controller.CloseSession(); } catch { }
        }

        /// <summary>拖动进行中（Esc 钩子据此放行按键给 DragMove 取消拖动）。</summary>
        public bool IsDragging { get { return _dragging; } }

        /// <summary>归属 Word 活动窗口：随 Word 最小化/置底，不作为全局置顶遮挡其他程序。</summary>
        private void TryAttachToWordWindow()
        {
            try
            {
                IntPtr hwnd = WordWindow.Get(_controller.App);
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
                BtnBold.IsChecked = false;
                BtnItalic.IsChecked = false;
                BtnUnderline.IsChecked = false;
                BtnStrike.IsChecked = false;
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

        /// <summary>会话样本信息 → 字号提示（样本单元首锚字号；NaN 未知显示 —）。</summary>
        public void UpdateSampleInfo(FormatSignature sampleSig)
        {
            _currentSizePt = sampleSig.SizePt;
            RefreshSizeDisplay();
        }

        private void RefreshSizeDisplay()
        {
            string t = _currentSizePt.HasValue ? _currentSizePt.Value.ToString("0.#") : "—";
            var block = FindName("TxtSize") as TextBlock;
            if (block != null && block.Text != t) block.Text = t;
        }

        /// <summary>应用字号并把提示刷成新值。</summary>
        private void ApplySize(double pt)
        {
            _currentSizePt = pt;
            RefreshSizeDisplay();
            _controller.ApplyChanges(new FormatChanges { SizePt = pt });
        }

        private void OnSizeUp(object sender, RoutedEventArgs e)
        {
            StepSize(+1);
        }

        private void OnSizeDown(object sender, RoutedEventArgs e)
        {
            StepSize(-1);
        }

        /// <summary>沿 Word 字号表步进；表外值先吸附到相邻表项，超出表范围按 ±1 续步（夹在 [1,1638]）。</summary>
        private void StepSize(int dir)
        {
            double cur = _currentSizePt ?? 10.5;
            double next = double.NaN;
            if (dir > 0)
            {
                foreach (double s in SizeSteps)
                    if (s > cur + 0.01) { next = s; break; }
                if (double.IsNaN(next)) next = Math.Min(1638, Math.Floor(cur) + 1);
            }
            else
            {
                for (int i = SizeSteps.Length - 1; i >= 0; i--)
                    if (SizeSteps[i] < cur - 0.01) { next = SizeSteps[i]; break; }
                if (double.IsNaN(next)) next = Math.Max(1, Math.Ceiling(cur) - 1);
            }
            if (!double.IsNaN(next) && Math.Abs(next - cur) >= 0.01)
                ApplySize(next);
        }

        /// <summary>字号下拉（与颜色面板同款弹层，点选即应用）。</summary>
        private void OnSizePopup(object sender, RoutedEventArgs e)
        {
            var panelBg = BrushOf("PopupBgBrush");
            var panelBorder = BrushOf("PopupBorderBrush");
            var accent = BrushOf("AccentBrush");
            var textBrush = BrushOf("TextBrush");
            var hoverBrush = BrushOf("ItemHoverBrush");

            var grid = new UniformGrid { Columns = 5 };
            foreach (double s in SizeSteps)
            {
                double val = s;
                var item = new Border
                {
                    Background = Brushes.Transparent, CornerRadius = new CornerRadius(3),
                    Margin = new Thickness(2), Padding = new Thickness(8, 5, 8, 5), Cursor = Cursors.Hand,
                    Child = new TextBlock
                    {
                        Text = val.ToString("0.#"), FontSize = 12.5,
                        Foreground = _currentSizePt.HasValue && Math.Abs(_currentSizePt.Value - val) < 0.01
                            ? accent : textBrush
                    }
                };
                if (_currentSizePt.HasValue && Math.Abs(_currentSizePt.Value - val) < 0.01)
                    item.Background = hoverBrush;
                item.MouseLeftButtonUp += (s2, _) => { if (_sizePopup != null) _sizePopup.IsOpen = false; ApplySize(val); };
                grid.Children.Add(item);
            }

            _sizePopup = new Popup
            {
                Child = new Border
                {
                    Background = panelBg, CornerRadius = new CornerRadius(3),
                    BorderBrush = panelBorder, BorderThickness = new Thickness(1),
                    Padding = new Thickness(5), Child = grid
                },
                PlacementTarget = BtnSize, Placement = PlacementMode.Bottom,
                StaysOpen = false, AllowsTransparency = true
            };
            _sizePopup.IsOpen = true;
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

        private void OnBoldToggled(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            // 二态开关：勾上=加粗，取消=取消加粗（旧版取消时静默不生效）
            bool bold = BtnBold.IsChecked == true;
            _controller.ApplyChanges(new FormatChanges { Bold = bold });
        }

        private void OnItalicToggled(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            _controller.ApplyChanges(new FormatChanges { Italic = BtnItalic.IsChecked == true });
        }

        private void OnUnderlineToggled(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            _controller.ApplyChanges(new FormatChanges { Underline = BtnUnderline.IsChecked == true });
        }

        private void OnStrikeToggled(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            _controller.ApplyChanges(new FormatChanges { StrikeThrough = BtnStrike.IsChecked == true });
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

        private void OnJumpNext(object sender, RoutedEventArgs e)
        {
            int idx = _controller.JumpToNextMatch();
            if (idx > 0)
            {
                int total = _controller.GetMatchedCount();
                BtnJump.Content = string.Format("下一处 {0}/{1} ▼", idx, total);
            }
        }

        /// <summary>强调色板：柔和低饱和系，与命中黄色基底/白纸/黑字协调。
        /// 审美取舍：排除红（警报感）、深蓝/墨绿/深紫（黑字辨识度差）、荧光粉绿（与黄互斥）。</summary>
        private static readonly string[] EmphasisHexes = { "#E8B34B", "#9DC3E6", "#A9D18E", "#C9A0DC" };
        private static readonly string[] EmphasisNames = { "蜜橙", "天蓝", "薄荷", "淡紫" };

        private void UpdateEmphasisSwatch()
        {
            try
            {
                SwEmphasis.Background = new SolidColorBrush(
                    (Color)ColorConverter.ConvertFromString(_controller.Settings.EmphasisColor ?? "#E8B34B"));
            }
            catch { }
        }

        private bool _emphasisFilled;

        private void OnEmphasisClick(object sender, RoutedEventArgs e)
        {
            UpdateEmphasisSwatch();
            if (!_emphasisFilled)
            {
                _emphasisFilled = true;
                var host = new StackPanel { Orientation = Orientation.Horizontal };
                for (int i = 0; i < EmphasisHexes.Length; i++)
                {
                    string hex = EmphasisHexes[i];
                    var item = new Button
                    {
                        Style = (Style)FindResource("TyPlain"),
                        ToolTip = EmphasisNames[i],
                        Padding = new Thickness(5), Margin = new Thickness(2), Cursor = Cursors.Hand
                    };
                    var dot = new Border
                    {
                        CornerRadius = new CornerRadius(3), Width = 20, Height = 14,
                        Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex)),
                        BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Color.FromArgb(0x80, 0, 0, 0))
                    };
                    item.Content = dot;
                    item.Click += (s2, e2) =>
                    {
                        _controller.SetEmphasisColor(hex);
                        UpdateEmphasisSwatch();
                        PopEmphasis.IsOpen = false;
                    };
                    host.Children.Add(item);
                }
                var panel = new StackPanel();
                panel.Children.Add(new TextBlock
                {
                    Text = "强调色（当前跳转处底色）", FontSize = 10.5, Margin = new Thickness(2, 0, 0, 4),
                    Foreground = (Brush)FindResource("TextSecondaryBrush")
                });
                panel.Children.Add(host);
                PopEmphasis.Child = new Border
                {
                    Background = (Brush)FindResource("PanelBrush"),
                    BorderBrush = (Brush)FindResource("PanelBorderBrush"),
                    BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(10)
                };
                ((Border)PopEmphasis.Child).Child = panel;
            }
            PopEmphasis.PlacementTarget = BtnEmphasis;
            PopEmphasis.IsOpen = true;
        }

        private void OnClose(object sender, RoutedEventArgs e)
        {
            _controller.CloseSession();
        }
    }
}
