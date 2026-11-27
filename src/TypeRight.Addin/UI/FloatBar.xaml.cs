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
using TypeRight.Formatting;
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
        private OfficeThemeKind _themeKind;
        private int? _fontColorRgb;   // 当前字体色 0xRRGGBB（null=本次会话未设置，取色器回退黑）

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

                // GetPoint 只对"当前在视口里"的段落返回屏幕坐标；段落滚出可见区时
                // 它吐的是负的文档坐标（实测 -31493,-30951），据此定位会把浮条丢到屏幕外。
                // 这种情形退回"Word 窗口中央偏下"——用户看得见，且样本本来也不在视口里。
                if (!LooksLikeScreenPoint(x, y))
                {
                    CenterOverWordWindow(app);
                    return;
                }

                double px = x / scale, py = y / scale, ph = h / scale;

                // 尺寸用**窗口实际物理尺寸**，不用 Width/ActualWidth。
                // 坑：Width=760 是含 Margin 的窗口宽，而内容实际 748；在 125% 下
                // 这点差异会累积成定位偏差（实测按钮命中点与视觉位置差 17px）。
                double barW, barH;
                try
                {
                    var r0 = new Win32Rect();
                    if (GetWindowRect(new WindowInteropHelper(this).Handle, out r0) && r0.Right > r0.Left)
                    {
                        barW = (r0.Right - r0.Left) / scale;
                        barH = (r0.Bottom - r0.Top) / scale;
                    }
                    else
                    {
                        barW = ActualWidth > 0 ? ActualWidth : Width;
                        barH = ActualHeight > 0 ? ActualHeight : 150;
                    }
                }
                catch
                {
                    barW = ActualWidth > 0 ? ActualWidth : Width;
                    barH = ActualHeight > 0 ? ActualHeight : 150;
                }

                // 显示器边界（全用**物理像素**判定，最后再换算回 WPF）。
                // 坑：这块机器的第二屏在 y=-900（负坐标），若把 WPF 坐标和物理坐标混着比，
                // 夹取基准整块错位——浮条会被夹到第二屏去。
                var screen = System.Windows.Forms.Screen.FromPoint(new System.Drawing.Point(x, y));
                System.Drawing.Rectangle pb = screen.Bounds;   // 物理像素

                // 目标左上角（物理像素）
                double physLeft = x + 18 * scale;
                double physTop = y + h + 8 * scale;

                // 下溢出 → 放到锚点上方
                if (physTop + barH * scale > pb.Bottom - 8 * scale)
                    physTop = y - barH * scale - 8 * scale;

                // 四边夹回该显示器（物理坐标）
                double maxLeft = pb.Right - barW * scale - 12 * scale;
                double minLeft = pb.Left + 8 * scale;
                if (physLeft > maxLeft) physLeft = maxLeft;
                if (physLeft < minLeft) physLeft = minLeft;

                double maxTop = pb.Bottom - barH * scale - 12 * scale;
                double minTop = pb.Top + 8 * scale;
                if (physTop > maxTop) physTop = maxTop;
                if (physTop < minTop) physTop = minTop;

                // 物理 → WPF（实测映射：物理 = WPF × scale）
                Left = physLeft / scale;
                Top = physTop / scale;
            }
            catch
            {
                try { CenterOverWordWindow(app); }
                catch { WindowStartupLocation = WindowStartupLocation.CenterScreen; }
            }
        }

        /// <summary>判断 GetPoint 返回的是不是真的屏幕坐标。
        /// 视口外时 Word 给的是负的文档坐标（量级几万），任何真实屏幕坐标都远不到那里。</summary>
        private static bool LooksLikeScreenPoint(int x, int y)
        {
            return x > -20000 && y > -20000 && x < 100000 && y < 100000;
        }

        /// <summary>兜底：Word 活动窗口中央偏下（不是屏幕中央——多显示器下更贴近视线）。
        /// 全程用**物理像素**计算，最后统一 ÷scale 转回 WPF，避免与负坐标显示器混算出错。</summary>
        private void CenterOverWordWindow(WdApplication app)
        {
            IntPtr hwnd = WordWindow.Get(app);
            if (hwnd == IntPtr.Zero) { WindowStartupLocation = WindowStartupLocation.CenterScreen; return; }
            var r = new Win32Rect();
            if (!GetWindowRect(hwnd, out r)) { WindowStartupLocation = WindowStartupLocation.CenterScreen; return; }
            double scale = GetDpiScale();
            double barW = ActualWidth > 0 ? ActualWidth : Width;
            double barH = ActualHeight > 0 ? ActualHeight : 150;

            // 用 Word 窗口所在显示器做夹取基准（物理像素）
            var screen = System.Windows.Forms.Screen.FromRectangle(
                new System.Drawing.Rectangle(r.Left, r.Top, Math.Max(1, r.Right - r.Left), Math.Max(1, r.Bottom - r.Top)));
            System.Drawing.Rectangle pb = screen.Bounds;

            // 目标位置（物理像素）：窗口水平居中、垂直略偏下
            double physLeft = (r.Left + r.Right) / 2.0 - barW * scale / 2;
            double physTop = (r.Top + r.Bottom) / 2.0 - barH * scale / 2 + 80 * scale;

            // 四边夹回该显示器（物理坐标；负坐标屏也成立）
            double maxLeft = pb.Right - barW * scale - 12 * scale;
            double minLeft = pb.Left + 8 * scale;
            if (physLeft > maxLeft) physLeft = maxLeft;
            if (physLeft < minLeft) physLeft = minLeft;

            double maxTop = pb.Bottom - barH * scale - 12 * scale;
            double minTop = pb.Top + 8 * scale;
            if (physTop > maxTop) physTop = maxTop;
            if (physTop < minTop) physTop = minTop;

            Left = physLeft / scale;
            Top = physTop / scale;
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

        /// <summary>DPI 缩放（物理像素 / DIP）。
        /// 关键：只在真拿到 DPI 时缓存。PositionNear 在 Show() 之前调用，那时窗口还没有
        /// HwndSource，早先的实现会把兜底值 1.0 永久缓存——125% 缩放下 GetPoint 给的物理像素
        /// 就再也不会被换算，浮条被定位到屏幕外（实测 Top=918 / 屏高 864）。</summary>
        private double GetDpiScale()
        {
            if (_dpiScale > 0) return _dpiScale;
            try
            {
                var source = System.Windows.Interop.HwndSource.FromVisual(this);
                if (source?.CompositionTarget != null)
                {
                    double s = source.CompositionTarget.TransformToDevice.M11;
                    if (s > 0) { _dpiScale = s; return _dpiScale; }
                }
            }
            catch { }

            // 窗口还没上屏：显式建 HWND 后向系统要 DPI（Win7 无此 API → 退回 GDI DpiX）
            try
            {
                IntPtr hwnd = new WindowInteropHelper(this).EnsureHandle();
                if (hwnd != IntPtr.Zero)
                {
                    if (TryGetDpiForWindow != null)
                    {
                        uint dpi = TryGetDpiForWindow(hwnd);
                        if (dpi > 0) { _dpiScale = dpi / 96.0; return _dpiScale; }
                    }
                    using (var g = System.Drawing.Graphics.FromHwnd(hwnd))
                    {
                        if (g.DpiX > 0) { _dpiScale = g.DpiX / 96.0; return _dpiScale; }
                    }
                }
            }
            catch { }

            // 都拿不到：本次按 1.0 算，但绝不缓存——下次调用还要再试（别把猜测固化成事实）
            return 1.0;
        }

        /// <summary>Win10 1607+ 的 GetDpiForWindow；老系统为 null（PowerShell/PInvoke 均在运行时尝试）。</summary>
        private static readonly Func<IntPtr, uint> TryGetDpiForWindow = LoadGetDpiForWindow();

        private static Func<IntPtr, uint> LoadGetDpiForWindow()
        {
            try
            {
                IntPtr user32 = GetModuleHandle("user32.dll");
                if (user32 == IntPtr.Zero) return null;
                IntPtr fn = GetProcAddress(user32, "GetDpiForWindow");
                if (fn == IntPtr.Zero) return null;
                return (Func<IntPtr, uint>)Marshal.GetDelegateForFunctionPointer(fn, typeof(Func<IntPtr, uint>));
            }
            catch { return null; }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

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
                        if (alive && bar.IsVisible && !bar.IsDragging && !bar.IsEditingText)
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

        /// <summary>Esc 钩子入口（拖动中留给 DragMove 自己取消，不重复处理）。
        /// 正在字体/字号框里打字时，Esc 的语义是"放弃这次编辑"，不是"关会话"——
        /// 由控件自己的 KeyDown 处理（并置 Handled），钩子这里必须放行。</summary>
        private void EscapeClose()
        {
            if (_dragging) return;
            try { _controller.CloseSession(); } catch { }
        }

        /// <summary>焦点是否在可编辑输入框内（字体/字号打字中）。</summary>
        private bool IsEditingText
        {
            get
            {
                try
                {
                    var focused = Keyboard.FocusedElement as FrameworkElement;
                    if (focused == null) return false;
                    return focused == CboFont
                        || focused == CboSize
                        || focused.Name == "PART_EditableTextBox";
                }
                catch { return false; }
            }
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
                CboFont.Text = string.Empty;
                BtnBold.IsChecked = false;
                BtnItalic.IsChecked = false;
                BtnUnderline.IsChecked = false;
                BtnStrike.IsChecked = false;
                ChkStrict.IsChecked = true;
                ChkIgnoreSize.IsChecked = false;
                ChkIgnoreColor.IsChecked = false;
                ChkIgnoreBold.IsChecked = false;
                ChkIgnoreCharClass.IsChecked = false;
                ChkIgnoreSize.IsEnabled = false;
                ChkIgnoreColor.IsEnabled = false;
                ChkIgnoreBold.IsEnabled = false;
                ChkIgnoreCharClass.IsEnabled = false;
                SetColorStrip(null);
                _fontColorRgb = null;
                // 字体/字号一起清：新会话在 UpdateSampleInfo 里按样本回填，
                // 避免上一个会话的字体字号残留到这一次（显示的在跑偏时最难排查）
                ClearFontDisplayText();
                _currentSizePt = null;
                RefreshSizeDisplay();
                CboFont.ToolTip = "当前字体（点击更换）";
            }
            finally { _suppressEvents = false; }
        }

        /// <summary>自动化取证：模拟在字号框打字并回车（验证手输链路真的通）。
        /// 返回最终显示快照；解析失败会有 toast，这里只关心最终显示的文本。</summary>
        public string AutomationTypeSize(string text)
        {
            try
            {
                CboSize.Text = text ?? "";
                CommitTypedSize();
                return GetDisplaySnapshot();
            }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        /// <summary>自动化取证：模拟在字体框打字并回车。</summary>
        public string AutomationTypeFont(string text)
        {
            try
            {
                CboFont.Text = text ?? "";
                CommitTypedFont();
                return GetDisplaySnapshot();
            }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        /// <summary>自动化取证：当前字体/字号显示的真实文本（"宋体|四号"）。
        /// 验收需要断言"浮条显示了什么"，而不是断言内部字段——显示错了字段也会是"对的"。</summary>
        public string GetDisplaySnapshot()
        {
            // 可编辑 ComboBox：手输后 SelectedItem 可能为 null 而 Text 有值，两者都要看
            string font = CboFont.Text;
            if (string.IsNullOrEmpty(font)) font = CboFont.SelectedItem as string;
            if (string.IsNullOrEmpty(font)) font = _currentFontName ?? "";
            return font + "|" + CurrentSizeText();
        }

        /// <summary>字号框当前文本（可编辑 ComboBox 的 Text；空则回退到"未读到"占位）。</summary>
        private string CurrentSizeText()
        {
            string t = CboSize.Text;
            return string.IsNullOrEmpty(t) ? "—" : t;
        }

        /// <summary>会话样本信息 → 字号提示（样本单元首锚字号；NaN 未知显示 —）。
        /// 同时回填样本字体：同类选择生效后浮条要像 Word「开始」选项卡那样
        /// 直接显示当前选中同类的字体与字号，而不是空白下拉 + 纯数字。
        /// 注意：文本同类模式下 SampleSignature 是 default(struct)，SizePt 为 0 而非 NaN，
        /// 必须按"非正数=未知"兜底，否则字号框会显示 0。</summary>
        public void UpdateSampleInfo(FormatSignature sampleSig)
        {
            _currentSizePt = IsUsableSize(sampleSig.SizePt) ? sampleSig.SizePt : (double?)null;
            RefreshSizeDisplay();
            RefreshSampleFontDisplay(sampleSig.EastAsiaFont);
        }

        /// <summary>合法字号判定：Word 字号恒为正；0/负数/NaN 都是"读不到"（default 结构体的 SizePt=0）。</summary>
        private static bool IsUsableSize(double pt)
        {
            return !double.IsNaN(pt) && !double.IsInfinity(pt) && pt > 0;
        }

        /// <summary>字号显示：中文字号档位显示「四号」/「小四」，非档位值仍显示数值。
        /// 文案口径统一走 ChineseFontSizes（纯逻辑层，xUnit 已覆盖）。
        /// 可编辑 ComboBox：把文本写回 Text（不改 Items，避免污染档位列表）。</summary>
        private void RefreshSizeDisplay()
        {
            string t = _currentSizePt.HasValue
                ? ChineseFontSizes.Label(_currentSizePt.Value)
                : "—";

            _suppressEvents = true;
            try
            {
                CboSize.Text = t;
                CboSize.ToolTip = _currentSizePt.HasValue
                    ? "当前 " + ChineseFontSizes.Description(_currentSizePt.Value) + "（可直接输入：四号 / 14 / 12.5）"
                    : "输入字号：四号、小四、14、12.5 或 14磅（回车确认）";
            }
            finally { _suppressEvents = false; }
        }

        /// <summary>字体框回填样本的中文字体（与字号同一时机）。
        /// 列表里没有该字体时（主题字体如「+中文正文」/第三方字体/已卸载字体）直接写 Text 显示原名，
        /// 而不是留空——Word 原生在同样情况下也显示这个名字，留空会让用户以为"没读到字体"。
        /// 用户手动改过字体后会话内不再回填，防把用户的输入冲掉。</summary>
        private void RefreshSampleFontDisplay(string eaFont)
        {
            if (_suppressEvents) return;
            if (string.IsNullOrEmpty(eaFont)) return;

            _suppressEvents = true;
            try
            {
                string matched = null;
                for (int i = 0; i < CboFont.Items.Count; i++)
                {
                    string item = CboFont.Items[i] as string;
                    if (string.Equals(item, eaFont, StringComparison.OrdinalIgnoreCase)) { matched = item; break; }
                }

                // 可编辑 ComboBox：有匹配就选中（下拉能高亮），否则直接写 Text 显示原名。
                // 不再往 Items 里插临时项——打字录入后插入/移除会把用户看得见的列表搞脏。
                if (matched != null) CboFont.SelectedItem = matched;
                else { CboFont.SelectedIndex = -1; CboFont.Text = eaFont; }

                _currentFontName = matched ?? eaFont;
                CboFont.ToolTip = "当前字体 " + eaFont + "（可直接输入更换）";
            }
            finally { _suppressEvents = false; }
        }

        /// <summary>清掉会话级的字体显示状态（换会话时调用）。</summary>
        private void ClearFontDisplayText()
        {
            _currentFontName = null;
            try { CboFont.Text = string.Empty; } catch { }
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

        /// <summary>字号框展开前填档位（中文名双行对照：四号 / 14）。
        /// 每次展开重建，保证"当前值"高亮跟着最新状态走。</summary>
        private void OnSizeDropDownOpened(object sender, EventArgs e)
        {
            _suppressEvents = true;
            try
            {
                CboSize.Items.Clear();
                foreach (double s in SizeSteps)
                    CboSize.Items.Add(new SizeChoice(s));
                CboSize.SelectedIndex = -1;   // 不留选择，"当前值"只体现在文字上
            }
            finally { _suppressEvents = false; }
        }

        /// <summary>字号档位项：显示「四号 / 14」。ToString 供可编辑 TextBox 采纳时取干净值。</summary>
        private sealed class SizeChoice
        {
            public readonly double Pt;
            public SizeChoice(double pt) { Pt = pt; }
            public override string ToString() { return ChineseFontSizes.Label(Pt); }
        }

        /// <summary>从下拉里选了一个档位（可编辑 ComboBox：用户打字时也会触发，需过滤）。</summary>
        private void OnSizeComboChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressEvents) return;
            var choice = CboSize.SelectedItem as SizeChoice;
            if (choice == null) return;   // 手输文本不是档位项 → 交给 KeyDown/失焦提交
            ApplySize(choice.Pt);
        }

        /// <summary>字号框按键：回车提交手输；Esc 放弃本次编辑（不关会话）；上下键交给下拉。</summary>
        private void OnSizeKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CommitTypedSize();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                // 正在改字号时 Esc：先还原显示，不关整个会话（用户意图是"不改了"）
                RefreshSizeDisplay();
                Keyboard.ClearFocus();
                e.Handled = true;
            }
        }

        /// <summary>提交手输字号。解析不了就还原显示并提示——绝不静默改成别的值。</summary>
        private void CommitTypedSize()
        {
            string raw = CboSize.Text;
            double? pt = ChineseFontSizes.TryParse(raw);
            if (!pt.HasValue)
            {
                RefreshSizeDisplay();
                _controller.ShowToast("字号「" + (raw ?? "").Trim() + "」无法识别：可输入 四号、小四、14、12.5 或 14磅");
                return;
            }
            ApplySize(pt.Value);
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

        /// <summary>字体框按键：回车提交手输字体名；Esc 放弃本次编辑（不关会话）。</summary>
        private void OnFontKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CommitTypedFont();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                // 正在改字体时 Esc：还原显示，不关整个会话
                _suppressEvents = true;
                try { if (_currentFontName != null) CboFont.Text = _currentFontName; } catch { }
                finally { _suppressEvents = false; }
                Keyboard.ClearFocus();
                e.Handled = true;
            }
        }

        /// <summary>离开字体框时提交手输（Word 同款行为：改完点别处即生效）。</summary>
        private void OnFontLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (_suppressEvents) return;
            CommitTypedFont();
        }

        private string _currentFontName;

        /// <summary>提交手输字体名。空输入=不做改动；不允许"静默清空"。</summary>
        private void CommitTypedFont()
        {
            string raw = (CboFont.Text ?? "").Trim();
            if (raw.Length == 0) return;
            if (!LooksLikeFontName(raw))
            {
                // 还原显示并提示：字体名里出现控制字符多半是误触/输入法残留
                _suppressEvents = true;
                try { CboFont.Text = _currentFontName ?? string.Empty; } catch { }
                finally { _suppressEvents = false; }
                _controller.ShowToast("字体名「" + raw + "」无效");
                return;
            }
            if (string.Equals(raw, _currentFontName, StringComparison.OrdinalIgnoreCase)) return;

            _currentFontName = raw;
            _controller.ApplyChanges(new FormatChanges { EastAsiaFont = raw });

            // 可编辑 ComboBox 在外部改 Text 后会把 Text 与 SelectedItem 弄成不一致；
            // 这里显式回写一次，保证框里显示的就是刚应用的字体（否则快照/界面会读回旧值）。
            _suppressEvents = true;
            try
            {
                CboFont.SelectedIndex = -1;
                CboFont.Text = raw;
                CboFont.ToolTip = "当前字体 " + raw + "（可直接输入更换）";
            }
            finally { _suppressEvents = false; }
        }

        /// <summary>字体名基本合法性：非空、无控制字符、长度合理（Word 字体名上限 32 字符）。</summary>
        private static bool LooksLikeFontName(string s)
        {
            if (string.IsNullOrEmpty(s) || s.Length > 32) return false;
            foreach (char c in s)
                if (char.IsControl(c)) return false;
            return true;
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
            // 常用色：柔和低饱和（讲义/试卷排版语义色）
            var presets = new[]
            {
                new object[] { "黑色", "#000000" }, new object[] { "深蓝", "#1A365D" },
                new object[] { "中蓝", "#2B6CB0" }, new object[] { "深红", "#C0392B" },
                new object[] { "深绿", "#2F855A" }, new object[] { "灰", "#595959" }
            };
            // 标准色：与 Word"字体颜色"下拉的标准色一致
            var standards = new[]
            {
                new object[] { "深红", "#C00000" }, new object[] { "红色", "#FF0000" },
                new object[] { "橙色", "#FFC000" }, new object[] { "黄色", "#FFFF00" },
                new object[] { "浅绿", "#92D050" }, new object[] { "绿色", "#00B050" },
                new object[] { "浅蓝", "#00B0F0" }, new object[] { "蓝色", "#0070C0" },
                new object[] { "深蓝", "#002060" }, new object[] { "紫色", "#7030A0" }
            };

            // 调色板（主题化：面板底色/描边/悬停环）
            Brush popupBg = BrushOf("PopupBgBrush"), popupBorder = BrushOf("PopupBorderBrush"),
                 accent = BrushOf("AccentBrush"), panelBorder = BrushOf("PanelBorderBrush");
            var labelBrush = BrushOf("TextSecondaryBrush");
            var hoverBrush = BrushOf("ItemHoverBrush");
            var textBrush = BrushOf("TextBrush");
            var sepBrush = BrushOf("SeparatorBrush");

            var panel = new StackPanel();
            panel.Children.Add(MakePaletteLabel("常用色", labelBrush, topMargin: 0));
            panel.Children.Add(MakeSwatchRow(presets, accent, panelBorder));
            panel.Children.Add(MakePaletteLabel("标准色", labelBrush, topMargin: 3));
            panel.Children.Add(MakeSwatchRow(standards, accent, panelBorder));
            panel.Children.Add(new Separator { Background = sepBrush, Height = 1, Margin = new Thickness(2, 7, 2, 5) });
            panel.Children.Add(MakeMoreColorsButton(hoverBrush, textBrush));

            _colorPopup = new Popup
            {
                Child = new Border
                {
                    Background = popupBg, CornerRadius = new CornerRadius(6),
                    BorderBrush = popupBorder,
                    BorderThickness = new Thickness(1), Padding = new Thickness(6),
                    Child = panel
                },
                PlacementTarget = BtnColor, Placement = PlacementMode.Bottom,
                StaysOpen = false, AllowsTransparency = true
            };
            _colorPopup.IsOpen = true;
        }

        private static TextBlock MakePaletteLabel(string text, Brush brush, int topMargin)
        {
            return new TextBlock
            {
                Text = text, FontSize = 11, Foreground = brush,
                Margin = new Thickness(4, topMargin, 0, 2)
            };
        }

        private StackPanel MakeSwatchRow(object[][] colors, Brush accent, Brush panelBorder)
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal };
            foreach (var p in colors)
            {
                var swatch = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString((string)p[1])),
                    Width = 20, Height = 20, CornerRadius = new CornerRadius(4),
                    Margin = new Thickness(3), BorderThickness = new Thickness(1),
                    BorderBrush = panelBorder, Cursor = Cursors.Hand, ToolTip = (string)p[0]
                };
                swatch.MouseEnter += (s, _) => { swatch.BorderBrush = accent; swatch.BorderThickness = new Thickness(2); };
                swatch.MouseLeave += (s, _) => { swatch.BorderBrush = panelBorder; swatch.BorderThickness = new Thickness(1); };
                swatch.MouseLeftButtonUp += OnPickColor;
                row.Children.Add(swatch);
            }
            return row;
        }

        /// <summary>"其他颜色…"入口（Word 同款）：四分色图标 + 文本，整行可点。</summary>
        private Border MakeMoreColorsButton(Brush hoverBrush, Brush textBrush)
        {
            var icon = new UniformGrid { Rows = 2, Columns = 2, Width = 13, Height = 13, VerticalAlignment = VerticalAlignment.Center };
            foreach (string hex in new[] { "#E81123", "#FFB900", "#0078D4", "#107C10" })
                icon.Children.Add(new Border { Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex)) });

            var content = new StackPanel { Orientation = Orientation.Horizontal };
            content.Children.Add(icon);
            content.Children.Add(new TextBlock
            {
                Text = "其他颜色…", FontSize = 12, Foreground = textBrush,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(7, 0, 0, 0)
            });

            var btn = new Border
            {
                Background = Brushes.Transparent, CornerRadius = new CornerRadius(4),
                Padding = new Thickness(5, 5, 8, 5), Margin = new Thickness(1, 0, 1, 1),
                Cursor = Cursors.Hand, Child = content
            };
            btn.MouseEnter += (s, _) => btn.Background = hoverBrush;
            btn.MouseLeave += (s, _) => btn.Background = Brushes.Transparent;
            btn.MouseLeftButtonUp += OnMoreColors;
            return btn;
        }

        private void OnMoreColors(object sender, MouseButtonEventArgs e)
        {
            if (_colorPopup != null) _colorPopup.IsOpen = false;
            try
            {
                int cur = _fontColorRgb ?? 0;
                using (var dlg = new System.Windows.Forms.ColorDialog
                {
                    FullOpen = true,   // 与 Word"其他颜色"一致：直接展开含自定义色的完整对话框
                    Color = System.Drawing.Color.FromArgb((cur >> 16) & 0xFF, (cur >> 8) & 0xFF, cur & 0xFF),
                    CustomColors = _controller.Settings.CustomFontColors
                })
                {
                    IntPtr owner = IntPtr.Zero;
                    try { owner = WordWindow.Get(_controller.App); } catch { }
                    var result = dlg.ShowDialog(owner == IntPtr.Zero ? null : new Win32Window(owner));

                    // 自定义色无论确定/取消都记住（Word 同款行为），跨会话持久化
                    try
                    {
                        if (dlg.CustomColors != null && dlg.CustomColors.Length > 0)
                        {
                            _controller.Settings.CustomFontColors = dlg.CustomColors;
                            _controller.SaveSettings();
                        }
                    }
                    catch { }

                    if (result != System.Windows.Forms.DialogResult.OK) return;
                    int rgb = (dlg.Color.R << 16) | (dlg.Color.G << 8) | dlg.Color.B;
                    _fontColorRgb = rgb;
                    SetColorStrip(rgb);
                    _controller.ApplyChanges(new FormatChanges { ColorRgb = rgb });
                }
            }
            catch (Exception ex) { TypeRightAddin.Log("MoreColors: " + ex.Message); }
        }

        /// <summary>把宿主窗口句柄包成 WinForms 模态属主（系统取色器要压在宿主窗口之上）。</summary>
        private sealed class Win32Window : System.Windows.Forms.IWin32Window
        {
            private readonly IntPtr _handle;
            public Win32Window(IntPtr handle) { _handle = handle; }
            public IntPtr Handle { get { return _handle; } }
        }

        private void OnPickColor(object sender, MouseButtonEventArgs e)
        {
            var border = (Border)sender;
            if (_colorPopup != null) _colorPopup.IsOpen = false;
            string hex = (string)border.Tag ?? ColorToHex(border.Background as SolidColorBrush);
            int rgb = int.Parse(hex.Substring(1), System.Globalization.NumberStyles.HexNumber);
            _fontColorRgb = rgb;
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
            ChkIgnoreCharClass.IsEnabled = !strict;
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
            if (ChkIgnoreCharClass.IsChecked == true) d |= FormatSignature.IgnoreDims.CharClass;
            return d;
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                if (_dragging) { e.Handled = true; return; } // 拖动中 Esc 仅取消本次拖动（DragMove 弹回原位）
                if (IsEditingText) return;                   // 打字中 Esc = 放弃本次编辑（控件自己处理）
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

        private void OnEmphasisMarkToggled(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            _controller.ApplyChanges(new FormatChanges { EmphasisMark = BtnEmphasisMark.IsChecked == true });
        }

        private void OnNoNegToggled(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            _controller.SetTextExcludeNegation(ChkNoNeg.IsChecked == true);
        }

        /// <summary>文本同类模式：切匹配行（格式条件 → 文本+否定前缀），浮条上无格式签名可显示。</summary>
        public void SetTextMode(string pattern, int count, bool excludeNegation)
        {
            _suppressEvents = true;
            try
            {
                ChkStrict.Visibility = Visibility.Collapsed;
                ChkIgnoreSize.Visibility = Visibility.Collapsed;
                ChkIgnoreColor.Visibility = Visibility.Collapsed;
                ChkIgnoreBold.Visibility = Visibility.Collapsed;
                ChkIgnoreCharClass.Visibility = Visibility.Collapsed;
                TxtPattern.Visibility = Visibility.Visible;
                TxtPattern.Text = string.Format("「{0}」× {1}", pattern, count);
                ChkNoNeg.Visibility = Visibility.Visible;
                ChkNoNeg.IsChecked = excludeNegation;
            }
            finally { _suppressEvents = false; }
        }

        /// <summary>恢复格式模式匹配行（每次 ShowBar 时调用；文本模式在 SetTextMode 再切换）。</summary>
        public void ResetToFormatMode()
        {
            _suppressEvents = true;
            try
            {
                ChkStrict.Visibility = Visibility.Visible;
                ChkIgnoreSize.Visibility = Visibility.Visible;
                ChkIgnoreColor.Visibility = Visibility.Visible;
                ChkIgnoreBold.Visibility = Visibility.Visible;
                ChkIgnoreCharClass.Visibility = Visibility.Visible;
                TxtPattern.Visibility = Visibility.Collapsed;
                ChkNoNeg.Visibility = Visibility.Collapsed;
            }
            finally { _suppressEvents = false; }
        }

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
                        CornerRadius = new CornerRadius(4), Width = 20, Height = 14,
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

        /// <summary>自动化取证：✕ 按钮相对窗口的**物理**位置（用于标定命中区）。
        /// 不用 PointToScreen —— 实测它的返回值与实际命中点偏差 ~17px（DPI 虚拟化）。
        /// 改为用窗口物理矩形 + 按钮在 WPF 树内的相对位置（×scale）来算。</summary>
        public string ProbeCloseButtonPhysical()
        {
            try
            {
                var helper = new WindowInteropHelper(this);
                var r = new Win32Rect();
                if (!GetWindowRect(helper.Handle, out r)) return "err:GetWindowRect";
                double scale = GetDpiScale();

                // 按钮相对窗口根元素的位置（WPF 逻辑单位）
                var root = Content as UIElement;
                if (root == null) return "err:no-root";
                Point p = BtnClose.TransformToAncestor(root).Transform(new Point(0, 0));
                double bw = BtnClose.ActualWidth, bh = BtnClose.ActualHeight;

                int physL = r.Left + (int)Math.Round(p.X * scale);
                int physT = r.Top + (int)Math.Round(p.Y * scale);
                int physW = (int)Math.Round(bw * scale);
                int physH = (int)Math.Round(bh * scale);

                return string.Format("win=({0},{1}) {2}x{3} scale={4:0.##} | btnWpf=({5:0.#},{6:0.#}) {7:0.#}x{8:0.#} | btnPhys=({9},{10}) {11}x{12} | center=({13},{14})",
                    r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, scale,
                    p.X, p.Y, bw, bh,
                    physL, physT, physW, physH,
                    physL + physW / 2, physT + physH / 2);
            }
            catch (Exception ex) { return "err:" + ex.Message; }
        }

        /// <summary>自动化取证：把浮条窗口移到指定 WPF 坐标，读回物理矩形，
        /// 用于标定 Left/Top(WPF) ↔ GetWindowRect(物理) 的真实映射关系。</summary>
        public string ProbeCoordinateMapping(double left, double top)
        {
            try
            {
                Left = left;
                Top = top;
                UpdateLayout();
                var helper = new WindowInteropHelper(this);
                IntPtr hwnd = helper.Handle;
                var r = new Win32Rect();
                bool ok = GetWindowRect(hwnd, out r);
                return string.Format("set=({0},{1}) phy=({2},{3}) size={4}x{5} scale={6:0.###} ok={7}",
                    left, top, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, GetDpiScale(), ok ? 1 : 0);
            }
            catch (Exception ex) { return "err:" + ex.Message; }
        }

        private void OnClose(object sender, RoutedEventArgs e)
        {
            TypeRightAddin.Log("FloatBar.OnClose: 点击了关闭按钮");
            _controller.CloseSession();
            TypeRightAddin.Log("FloatBar.OnClose: CloseSession 返回, IsVisible=" + IsVisible);
        }

        /// <summary>自动化取证：读当前窗口是否可见 + 关闭按钮的屏幕坐标（用于真机点击测试）。</summary>
        public string GetCloseButtonProbe()
        {
            try
            {
                if (!IsLoaded || !IsVisible) return "not-visible";
                Point p = BtnClose.PointToScreen(new Point(BtnClose.ActualWidth / 2, BtnClose.ActualHeight / 2));
                return string.Format("visible=1 btnCenter=({0},{1}) btnSize={2}x{3} barVisible={4}",
                    (int)p.X, (int)p.Y, (int)BtnClose.ActualWidth, (int)BtnClose.ActualHeight, IsVisible ? 1 : 0);
            }
            catch (Exception ex) { return "err:" + ex.Message; }
        }
    }
}
