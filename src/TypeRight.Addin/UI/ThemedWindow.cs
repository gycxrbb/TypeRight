using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace TypeRight.Addin.UI
{
    /// <summary>
    /// 宿主主题化的窗口基类 —— 本插件所有独立窗口统一从这里取样式。
    ///
    /// 设计约束（PRODUCT.md / DESIGN.md）：插件是 Word/WPS 的延伸，观感必须跟随宿主主题，
    /// 绝不能出现"宿主深色、插件白底浅字"这种看得见但看不清的组合。
    ///
    /// 实现要点：
    ///   · 颜色唯一真值 = ThemePalette（与 FloatBar 同源），窗口一律消费 P.* / 资源键
    ///   · 所有前景/背景走 **DynamicResource**（键名与 FloatBar.xaml 一致），
    ///     宿主切主题时就地刷新，不需要重建窗口
    ///   · 窗口背景取 P.Panel —— 2026-11 事故根因正是背景写死 Brushes.White
    ///     而文字取深色主题的浅色值，标题直接"消失"
    /// </summary>
    public abstract class ThemedWindow : Window
    {
        internal ThemePalette P { get; private set; }

        protected OfficeThemeKind ThemeKind { get; private set; }

        protected ThemedWindow()
        {
            ThemeKind = OfficeThemeKind.Light;
            P = ThemePalette.Get(ThemeKind);

            FontFamily = new FontFamily("Segoe UI, Microsoft YaHei UI");
            FontSize = 12.5;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ShowInTaskbar = false;
            UseLayoutRounding = true;
            SnapsToDevicePixels = true;

            PublishTokens();

            // 窗口底色也必须走资源引用。
            // 2026-11 事故：这里曾经什么都不设 → Window.Background 保持 WPF 默认纯白，
            // 而文字/卡片用的是深色主题令牌 → 浅字落白底，对比度只有 2.0，标题看不见。
            // 注意：不能写成 `Background = new SolidColorBrush(P.Panel)` ——那是构造时的
            // Light 快照，之后 ApplyTheme 换令牌也不会跟着变。
            UseResource(this, Control.BackgroundProperty, "PanelBrush");

            // 文字色必须走 TextElement.Foreground（可继承属性），而不是 Window.Foreground。
            // 2026-11 事故：只设 Window.Foreground 时，子控件里 new TextBlock{} 不带 Foreground
            // 会落到 WPF 默认的 #1F1F1F 黑字 → 深色底上对比度仅 1.17，等于看不见。
            UseResource(this, TextElement.ForegroundProperty, "TextBrush");

            // 迟到的宿主主题：Loaded 时再采一次（此时 Word 窗口句柄一定就绪）
            Loaded += OnFirstLoaded;
        }

        private bool _themeResolved;

        private void OnFirstLoaded(object sender, RoutedEventArgs e)
        {
            if (_themeResolved) return;
            _themeResolved = true;
            try { ApplyTheme(OfficeTheme.Detect(ResolveApp())); }
            catch { }
        }

        /// <summary>宿主 Application；子类应重写交出真实实例（否则 OfficeTheme 走 FindWindow 兜底）。</summary>
        protected virtual NetOffice.WordApi.Application ResolveApp()
        {
            return null;
        }

        /// <summary>套用主题：刷新资源令牌，所有 DynamicResource 引用处即时跟随。</summary>
        public void ApplyTheme(OfficeThemeKind kind)
        {
            ThemeKind = kind;
            P = ThemePalette.Get(kind);
            PublishTokens();
            // 背景与文字都走资源引用（构造函数里已建立 SetResourceReference），
            // 刷新令牌后自动跟随。此处绝不能再 `Background = new SolidColorBrush(...)`：
            // 直接赋值会覆盖 SetResourceReference 建立的绑定，之后换主题就不再跟随。
            OnThemeApplied();
        }

        /// <summary>子类若手工设过静态色，在此重刷；用 DynamicResource 的不需要。</summary>
        protected virtual void OnThemeApplied() { }

        private void PublishTokens()
        {
            Set("PanelBrush", P.Panel);
            Set("PanelBorderBrush", P.PanelBorder);
            Set("TextBrush", P.Text);
            Set("TextSecondaryBrush", P.TextSecondary);
            Set("InputBgBrush", P.InputBg);
            Set("InputBorderBrush", P.InputBorder);
            Set("HoverBrush", P.Hover);
            Set("PressedBrush", P.Pressed);
            Set("AccentBrush", P.Accent);
            Set("AccentHoverBrush", P.AccentHover);
            Set("PopupBgBrush", P.PopupBg);
            Set("PopupBorderBrush", P.PopupBorder);
            Set("ItemHoverBrush", P.ItemHover);
            Set("SeparatorBrush", P.Separator);
            Set("DisabledBrush", P.Disabled);
            Set("SuccessBrush", P.Success);
            Set("DangerBrush", P.Danger);
        }

        /// <summary>发布一个令牌。
        ///
        /// 坑：窗口可见后直接 `Resources[key] = b` **不会**让已经建立的
        /// DynamicResource 引用重新求值（WPF 对同键赋值不保证发变更通知）。
        /// 必须先 Remove 再 Add —— 移除本身会触发引用失效，随后的添加让它们取到新值。
        /// 2026-11 实测：不这么做的话 ApplyTheme(Dark) 之后窗口底色仍是 Light 的 #FAFAFA。</summary>
        private void Set(string key, Color c)
        {
            var b = new SolidColorBrush(c);
            b.Freeze();
            Resources.Remove(key);
            Resources.Add(key, b);
        }

        /// <summary>取当前色（仅在需要静态色时用；优先用 DynamicResource）。</summary>
        protected Brush B(string key)
        {
            return TryFindResource(key) as Brush ?? new SolidColorBrush(P.Text);
        }

        /// <summary>把控件某属性绑到主题资源上（运行中切主题自动跟随）。</summary>
        protected static void UseResource(FrameworkElement el, DependencyProperty prop, string key)
        {
            el.SetResourceReference(prop, key);
        }

        // ───────────────────── 样式工厂（所有窗口共用） ─────────────────────

        public TextBlock TitleText(string text)
        {
            var tb = new TextBlock { Text = text, FontSize = 17, FontWeight = FontWeights.SemiBold };
            UseResource(tb, TextBlock.ForegroundProperty, "TextBrush");
            return tb;
        }

        public TextBlock BodyText(string text, double size = 12.5)
        {
            var tb = new TextBlock { Text = text, FontSize = size, TextWrapping = TextWrapping.Wrap };
            UseResource(tb, TextBlock.ForegroundProperty, "TextBrush");
            return tb;
        }

        public TextBlock CaptionText(string text)
        {
            var tb = new TextBlock { Text = text, FontSize = 12, TextWrapping = TextWrapping.Wrap };
            UseResource(tb, TextBlock.ForegroundProperty, "TextSecondaryBrush");
            return tb;
        }

        public TextBox StyledTextBox(string text = null)
        {
            var t = new TextBox
            {
                Text = text ?? "",
                FontSize = 13,
                Padding = new Thickness(7, 5, 7, 5),
                BorderThickness = new Thickness(1)
            };
            UseResource(t, Control.ForegroundProperty, "TextBrush");
            UseResource(t, Control.BackgroundProperty, "InputBgBrush");
            UseResource(t, Control.BorderBrushProperty, "InputBorderBrush");
            UseResource(t, TextBox.CaretBrushProperty, "TextBrush");
            return t;
        }

        public Button PrimaryButton(string text)
        {
            var b = new Button
            {
                Content = text,
                Padding = new Thickness(18, 5, 18, 5),
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,      // Accent 实底上一律白字，两主题都成立
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            UseResource(b, Control.BackgroundProperty, "AccentBrush");
            return b;
        }

        public Button SecondaryButton(string text)
        {
            var b = new Button
            {
                Content = text,
                Padding = new Thickness(12, 4, 12, 4),
                FontSize = 12,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand
            };
            UseResource(b, Control.ForegroundProperty, "TextSecondaryBrush");
            UseResource(b, Control.BorderBrushProperty, "PanelBorderBrush");
            return b;
        }

        public CheckBox StyledCheckBox(string text, bool isChecked)
        {
            var c = new CheckBox
            {
                Content = text,
                FontSize = 12.5,
                IsChecked = isChecked,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            UseResource(c, Control.ForegroundProperty, "TextBrush");
            return c;
        }

        /// <summary>卡片容器：底与边走资源引用，宿主切主题自动跟随。</summary>
        public Border Card(UIElement child, double radius = 8)
        {
            var border = new Border
            {
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(radius),
                Padding = new Thickness(14),
                Child = child
            };
            UseResource(border, Border.BackgroundProperty, "InputBgBrush");
            UseResource(border, Border.BorderBrushProperty, "PanelBorderBrush");
            return border;
        }
    }
}
