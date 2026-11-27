using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using Microsoft.Win32;
using TypeRight.Addin.Controllers;
using TypeRight.Profiles;

namespace TypeRight.Addin.UI
{
    /// <summary>
    /// 「我的方案」窗口 —— 只用和管理，不负责存。
    ///
    /// 语境不同就不该共用一个窗口：用户点「方案」时想的是"挑一个用到这份文档"，
    /// 点「存为方案」时想的是"把刚才那步记下来"。混在一屏，两件事都做不干净
    /// （这正是旧面板的问题：标题写着"存为方案"，列表里却是"应用到当前文档"）。
    ///
    /// 这里只做三件事：
    ///   1. 列表：每个方案用人话说清"它会把什么样的段落改成什么样"
    ///   2. 应用：点一下套到当前文档
    ///   3. 低频管理：重命名 / 删除（导入导出收进"更多"，因为它们绝大多数用户用不到）
    /// </summary>
    public sealed class ProfileManagerWindow : ThemedWindow
    {
        private readonly TypeRightController _controller;
        private ProfileStoreV2 _store;
        private ListBox _list;
        private TextBlock _status;
        private TextBlock _empty;

        public ProfileManagerWindow(TypeRightController controller)
        {
            _controller = controller;
            Title = "我的方案";
            Width = 640;
            Height = 560;
            MinWidth = 520;
            MinHeight = 380;

            Build();
            Loaded += (s, e) => Reload();
        }

        /// <summary>主题跟随宿主 Word，必须把真实 Application 交出去。</summary>
        protected override NetOffice.WordApi.Application ResolveApp()
        {
            try { return _controller.App; } catch { return null; }
        }

        private void Build()
        {
            var root = new DockPanel { Margin = new Thickness(20, 18, 20, 16) };

            var head = new StackPanel { Margin = new Thickness(0, 0, 0, 14) };
            head.Children.Add(TitleText("我的方案"));
            var sub = CaptionText("选一个应用到当前文档，它会自动把同类段落改成方案里的格式。");
            sub.Margin = new Thickness(0, 4, 0, 0);
            head.Children.Add(sub);
            DockPanel.SetDock(head, Dock.Top);
            root.Children.Add(head);

            // 底部：状态 + 管理动作
            var bottom = new StackPanel();
            DockPanel.SetDock(bottom, Dock.Bottom);

            var tools = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 8)
            };
            var rename = SecondaryButton("重命名");
            rename.Click += OnRename;
            rename.Margin = new Thickness(0, 0, 6, 0);
            tools.Children.Add(rename);

            var del = SecondaryButton("删除");
            del.Click += OnDelete;
            del.Margin = new Thickness(0, 0, 14, 0);
            tools.Children.Add(del);

            // 导入导出：绝大多数人不需要，收成两个小字链接，不占视觉
            var import = LinkButton("导入方案文件…", OnImport);
            tools.Children.Add(import);
            var export = LinkButton("导出…", OnExport);
            export.Margin = new Thickness(8, 0, 0, 0);
            tools.Children.Add(export);

            DockPanel.SetDock(tools, Dock.Top);
            bottom.Children.Add(tools);

            _status = CaptionText("");
            bottom.Children.Add(_status);
            root.Children.Add(bottom);

            // 中部：列表 / 空态
            var center = new Grid();

            _list = new ListBox
            {
                FontSize = 13,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                ItemTemplate = CreateCardTemplate()
            };
            ScrollViewer.SetHorizontalScrollBarVisibility(_list, ScrollBarVisibility.Disabled);
            _list.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(OnApply));
            center.Children.Add(_list);

            _empty = new TextBlock
            {
                FontSize = 12.5,
                TextWrapping = TextWrapping.Wrap,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center,
                Visibility = Visibility.Collapsed,
                Text = "还没有方案。\n\n先在文档里用「同类选择」调整一次格式，关掉浮条后点提示里的「存为方案」，\n以后就能一键还原同样的操作。"
            };
            UseResource(_empty, TextBlock.ForegroundProperty, "TextSecondaryBrush");
            center.Children.Add(_empty);

            root.Children.Add(center);
            Content = root;
        }

        /// <summary>卡片模板：只展示"挑中什么 → 改成什么"，不出现 True / 规则条数 / 应用按钮以外的杂物。</summary>
        private DataTemplate CreateCardTemplate()
        {
            // 松散 XAML（无 x:Class）：绝不可写 Click="..."，否则布局阶段 XamlParseException 直接带崩 Word。
            // 点击统一由 ListBox 上的 ButtonBase.ClickEvent 冒泡处理。
            //
            // 颜色一律 DynamicResource 引主题令牌（不能写死 #FAFAFA / #1F1F1F）：
            // 宿主深色主题下写死浅底浅字 = 卡片整片看不见。这些键由 ThemedWindow 在
            // Resources 里发布，模板在框架内解析时能顺着逻辑树找到。
            string xaml = @"
<DataTemplate xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation"">
    <Border Padding=""14,12"" Margin=""0,0,0,8"" CornerRadius=""8""
            Background=""{DynamicResource InputBgBrush}""
            BorderBrush=""{DynamicResource PanelBorderBrush}"" BorderThickness=""1"">
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width=""*""/>
                <ColumnDefinition Width=""Auto""/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text=""{Binding Name}"" FontSize=""14.5"" FontWeight=""SemiBold""
                           Foreground=""{DynamicResource TextBrush}""/>
                <ItemsControl ItemsSource=""{Binding Rules}"" Margin=""0,7,0,0"">
                    <ItemsControl.ItemTemplate>
                        <DataTemplate>
                            <StackPanel Margin=""0,0,0,7"">
                                <TextBlock Text=""{Binding MatchText}"" FontSize=""12""
                                           Foreground=""{DynamicResource TextSecondaryBrush}""
                                           TextWrapping=""Wrap""/>
                                <TextBlock Text=""{Binding ChangeText}"" FontSize=""12.5"" FontWeight=""SemiBold""
                                           Foreground=""{DynamicResource AccentBrush}""
                                           TextWrapping=""Wrap"" Margin=""0,2,0,0""/>
                            </StackPanel>
                        </DataTemplate>
                    </ItemsControl.ItemTemplate>
                </ItemsControl>
            </StackPanel>
            <Button Grid.Column=""1"" Content=""应用到当前文档"" Padding=""13,5""
                    VerticalAlignment=""Center"" FontSize=""12.5"" Tag=""{Binding}""
                    Background=""{DynamicResource InputBgBrush}""
                    Foreground=""{DynamicResource TextBrush}""
                    BorderBrush=""{DynamicResource PanelBorderBrush}"" BorderThickness=""1""/>
        </Grid>
    </Border>
</DataTemplate>";
            return (DataTemplate)System.Windows.Markup.XamlReader.Parse(xaml);
        }

        private Button LinkButton(string text, RoutedEventHandler onClick)
        {
            var b = new Button
            {
                Content = text,
                FontSize = 11.5,
                Padding = new Thickness(2, 0, 2, 0),
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Center
            };
            UseResource(b, Control.ForegroundProperty, "TextSecondaryBrush");
            b.Click += onClick;
            return b;
        }

        private void Reload()
        {
            try
            {
                var settings = SettingsStore.Load();
                _store = new ProfileStoreV2(settings.ProfilesDirectory);
                var errors = new List<string>();
                var profiles = _store.LoadAll(errors);

                var cards = new List<ProfileCard>();
                foreach (ProfileV2 p in profiles) cards.Add(ProfileCard.From(p));
                _list.ItemsSource = cards;

                bool empty = cards.Count == 0;
                _empty.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
                _list.Visibility = empty ? Visibility.Collapsed : Visibility.Visible;

                _status.Text = errors.Count > 0
                    ? string.Join("；", errors)
                    : (empty ? "" : "共 " + cards.Count + " 个方案");
                _status.Foreground = errors.Count > 0 ? B("DangerBrush") : B("TextSecondaryBrush");
            }
            catch (Exception ex)
            {
                _status.Text = ex.Message;
                _status.Foreground = B("DangerBrush");
            }
        }

        private ProfileV2 Selected()
        {
            var card = _list.SelectedItem as ProfileCard;
            return card != null ? card.Profile : null;
        }

        private void OnApply(object sender, RoutedEventArgs e)
        {
            var btn = e.OriginalSource as Button;
            if (btn == null) return;
            var card = btn.Tag as ProfileCard;
            if (card == null) return;
            Close();
            _controller.ApplyProfile(card.Profile);
        }

        private void OnRename(object sender, RoutedEventArgs e)
        {
            var p = Selected();
            if (p == null) { _status.Text = "先在列表里点一下要重命名的方案。"; return; }
            var input = new InputDialog("重命名方案", p.Name) { Owner = this };
            if (input.ShowDialog() == true && input.Value.Trim().Length > 0)
            {
                _store.Rename(p, input.Value.Trim());
                Reload();
            }
        }

        private void OnDelete(object sender, RoutedEventArgs e)
        {
            var p = Selected();
            if (p == null) { _status.Text = "先在列表里点一下要删除的方案。"; return; }
            if (MessageBox.Show(this, "删除方案「" + p.Name + "」？\r\n（只删这份方案，不影响任何文档）",
                    "确认删除", MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK) return;
            _store.Delete(p.Id);
            Reload();
        }

        /// <summary>导出：把一个方案存成文件，主要用途是隔空传给同事/另一台电脑。</summary>
        private void OnExport(object sender, RoutedEventArgs e)
        {
            var p = Selected();
            if (p == null) { _status.Text = "先在列表里点一下要导出的方案。"; return; }
            var dlg = new SaveFileDialog
            {
                Title = "导出方案",
                Filter = "TypeRight 方案 (*.json)|*.json",
                FileName = p.Id + ".json"
            };
            if (dlg.ShowDialog(this) != true) return;
            _store.Export(p, dlg.FileName);
            _status.Text = "已导出到 " + System.IO.Path.GetFileName(dlg.FileName);
        }

        /// <summary>导入：把别人给的方案文件加进来。</summary>
        private void OnImport(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "导入方案",
                Filter = "TypeRight 方案 (*.json)|*.json"
            };
            if (dlg.ShowDialog(this) != true) return;
            List<string> errors;
            string id = _store.Import(dlg.FileName, out errors);
            if (id == null)
            {
                _status.Text = string.Join("；", errors);
                _status.Foreground = B("DangerBrush");
                return;
            }
            Reload();
            _status.Text = "已导入。";
        }
    }

    /// <summary>方案卡片视图模型。
    /// 必须是**属性**——WPF {Binding} 不认 public field，绑字段会静默空白。</summary>
    public sealed class ProfileCard
    {
        public ProfileV2 Profile { get; set; }
        public string Name { get; set; }
        public List<RuleSummary> Rules { get; set; }

        public static ProfileCard From(ProfileV2 p)
        {
            ProfileSummary s = ProfileDescriber.Describe(p);
            return new ProfileCard
            {
                Profile = p,
                Name = s.Name,
                Rules = s.Rules
            };
        }
    }

    /// <summary>简单输入对话框（重命名用）。继承 ThemedWindow 而不是裸 Window——
    /// 否则深色宿主主题下它是白底 + 深字，与插件其他窗口割裂。</summary>
    public sealed class InputDialog : ThemedWindow
    {
        private readonly TextBox _box;

        public string Value { get { return _box.Text; } }

        public InputDialog(string title, string initial)
        {
            Title = title;
            Width = 340;
            SizeToContent = SizeToContent.Height;
            ResizeMode = ResizeMode.NoResize;

            var panel = new StackPanel { Margin = new Thickness(16) };
            _box = StyledTextBox(initial ?? "");
            panel.Children.Add(_box);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            var ok = PrimaryButton("确定");
            ok.IsDefault = true;
            ok.Padding = new Thickness(16, 4, 16, 4);
            ok.Click += (s, e) => { DialogResult = true; };

            var cancel = SecondaryButton("取消");
            cancel.IsCancel = true;
            cancel.Margin = new Thickness(8, 0, 0, 0);
            cancel.Click += (s, e) => { DialogResult = false; };

            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            panel.Children.Add(buttons);

            Content = panel;
            Loaded += (s, e) => { _box.SelectAll(); _box.Focus(); };
        }
    }
}
