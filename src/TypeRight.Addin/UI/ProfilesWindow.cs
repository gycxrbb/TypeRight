using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using TypeRight.Addin.Controllers;
using TypeRight.Profiles;

namespace TypeRight.Addin.UI
{
    /// <summary>
    /// 方案面板（PRD v2 P1-1）：列表 + 应用；工具：存为方案（录制本次操作）/ 导入导出 / 重命名删除。
    /// </summary>
    public partial class ProfilesWindow : Window
    {
        private readonly TypeRightController _controller;
        private ProfileStoreV2 _store;

        public ProfilesWindow(TypeRightController controller, bool saveMode)
        {
            _controller = controller;
            Title = "TypeRight 方案";
            Width = 560; Height = 480;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ShowInTaskbar = false;

            BuildContent(saveMode);
            Loaded += (s, e) => Reload();
        }

        private ListBox _list;
        private TextBlock _status;
        private TextBlock _saveHeader;
        private TextBox _txtName;
        private CheckBox _chkCatchAll;
        private ComboBox _cboCatchAllFont;
        private CheckBox _chkConvert;

        private void BuildContent(bool saveMode)
        {
            var root = new DockPanel { Margin = new Thickness(14) };

            // 顶部：存方案区（可折叠）
            _saveHeader = new TextBlock
            {
                Text = _controller.PendingOpCount >= 2
                    ? "把刚才 " + _controller.PendingOpCount + " 步批量操作存为方案："
                    : "本次会话还没有可录制的批量操作（对同类段落改 2 次以上格式后可存方案）。",
                FontWeight = FontWeights.SemiBold,
                FontSize = 12.5,
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap
            };
            DockPanel.SetDock(_saveHeader, Dock.Top);
            root.Children.Add(_saveHeader);

            var savePanel = new StackPanel { Margin = new Thickness(0, 0, 0, 12), Visibility = Visibility.Collapsed };
            DockPanel.SetDock(savePanel, Dock.Top);

            var row1 = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 6) };
            row1.Children.Add(new TextBlock { Text = "方案名：", FontSize = 12, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0) });
            _txtName = new TextBox { Width = 200, Padding = new Thickness(4, 2, 4, 2) };
            row1.Children.Add(_txtName);
            savePanel.Children.Add(row1);

            var row2 = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 6) };
            _chkCatchAll = new CheckBox
            {
                Content = "其余段落（未命中任何规则的）统一为",
                FontSize = 12,
                VerticalAlignment = VerticalAlignment.Center
            };
            row2.Children.Add(_chkCatchAll);
            _cboCatchAllFont = new ComboBox { Width = 110, Margin = new Thickness(6, 0, 0, 0), FontSize = 12 };
            foreach (string f in new[] { "宋体", "仿宋", "楷体", "黑体", "微软雅黑" }) _cboCatchAllFont.Items.Add(f);
            _cboCatchAllFont.SelectedIndex = 0;
            row2.Children.Add(_cboCatchAllFont);
            savePanel.Children.Add(row2);

            var row3 = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 6) };
            _chkConvert = new CheckBox { Content = "应用时同时转换全文公式", FontSize = 12, IsChecked = true };
            row3.Children.Add(_chkConvert);
            var btnSave = new Button { Content = "存为方案", Padding = new Thickness(14, 3, 14, 3), Margin = new Thickness(12, 0, 0, 0) };
            btnSave.Click += OnSave;
            row3.Children.Add(btnSave);
            savePanel.Children.Add(row3);

            if (saveMode && _controller.PendingOpCount >= 2)
            {
                savePanel.Visibility = Visibility.Visible;
                _txtName.Text = "我的方案 " + DateTime.Now.ToString("MMdd");
            }
            else if (_controller.PendingOpCount >= 2)
            {
                savePanel.Visibility = Visibility.Visible;
            }
            root.Children.Add(savePanel);

            // 底部：状态 + 工具按钮
            var bottom = new DockPanel { Margin = new Thickness(0, 10, 0, 0) };
            DockPanel.SetDock(bottom, Dock.Bottom);

            var tools = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Left };
            AddTool(tools, "导入…", OnImport);
            AddTool(tools, "导出…", OnExport);
            AddTool(tools, "重命名", OnRename);
            AddTool(tools, "删除", OnDelete);
            bottom.Children.Add(tools);

            _status = new TextBlock { FontSize = 11, Foreground = new SolidColorBrush(Color.FromRgb(0x2F, 0x85, 0x5A)), HorizontalAlignment = HorizontalAlignment.Right };
            bottom.Children.Add(_status);

            // 中部：方案列表
            _list = new ListBox { FontSize = 13 };
            _list.ItemTemplate = CreateItemTemplate();
            root.Children.Add(_list);

            Content = root;
        }

        private static void AddTool(StackPanel parent, string label, RoutedEventHandler onClick)
        {
            var b = new Button { Content = label, Padding = new Thickness(8, 2, 8, 2), Margin = new Thickness(0, 0, 6, 0), FontSize = 12 };
            b.Click += onClick;
            parent.Children.Add(b);
        }

        private System.Windows.DataTemplate CreateItemTemplate()
        {
            string xaml = @"
<DataTemplate xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation"">
    <Border Padding=""8,6"" Margin=""0,2"" CornerRadius=""6"" Background=""#F5F7FB"">
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width=""*""/>
                <ColumnDefinition Width=""Auto""/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text=""{Binding Name}"" FontWeight=""SemiBold""/>
                <TextBlock FontSize=""11"" Foreground=""#6B7280"">
                    <Run Text=""{Binding Rules.Count, Mode=OneWay, StringFormat='规则 {0} 条'}""/><Run Text=""  ·  公式转换 ""/><Run Text=""{Binding ConvertFormulas, Mode=OneWay}""/>
                </TextBlock>
            </StackPanel>
            <Button Grid.Column=""1"" Content=""应用到当前文档"" Padding=""10,3"" Click=""OnApplyProfile"" Tag=""{Binding}""/>
        </Grid>
    </Border>
</DataTemplate>";
            var template = (System.Windows.DataTemplate)System.Windows.Markup.XamlReader.Parse(xaml);
            return template;
        }

        private void OnApplyProfile(object sender, RoutedEventArgs e)
        {
            var profile = (ProfileV2)((System.Windows.Controls.Button)sender).Tag;
            Close();
            _controller.ApplyProfile(profile);
        }

        private void Reload()
        {
            try
            {
                var settings = SettingsStore.Load();
                _store = new ProfileStoreV2(settings.ProfilesDirectory);
                var errors = new List<string>();
                var profiles = _store.LoadAll(errors);
                _list.ItemsSource = profiles;
                _status.Text = errors.Count > 0 ? string.Join("；", errors) : (profiles.Count + " 个方案");
                if (errors.Count > 0) _status.Foreground = new SolidColorBrush(Color.FromRgb(0xC0, 0x39, 0x2B));
            }
            catch (Exception ex)
            {
                _status.Text = ex.Message;
            }
        }

        private ProfileV2 Selected()
        {
            return _list.SelectedItem as ProfileV2;
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            string name = (_txtName.Text ?? "").Trim();
            if (name.Length == 0)
            {
                MessageBox.Show(this, "请给方案起个名字。");
                return;
            }
            try
            {
                SaveProfile(name);
            }
            catch (ProfileValidationException ex)
            {
                MessageBox.Show(this, string.Join("\r\n", ex.Errors), "校验未通过");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "保存失败");
            }
        }

        private void SaveProfile(string name)
        {
            string dir = _controller.SaveProfile(name,
                _chkCatchAll.IsChecked == true,
                _cboCatchAllFont.SelectedItem as string,
                _chkConvert.IsChecked == true);
            _status.Text = "已保存 → " + dir;
            _saveHeader.Text = "已保存为「" + name + "」，下一份同类文档打开后直接应用。";
            Reload();
        }

        private void OnImport(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Title = "导入方案 JSON", Filter = "TypeRight 方案 (*.json)|*.json" };
            if (dlg.ShowDialog(this) != true) return;
            List<string> errors;
            string id = _store.Import(dlg.FileName, out errors);
            if (id == null) MessageBox.Show(this, string.Join("\r\n", errors), "导入被拒绝");
            Reload();
        }

        private void OnExport(object sender, RoutedEventArgs e)
        {
            var p = Selected();
            if (p == null) { MessageBox.Show(this, "先在列表中选中一个方案。"); return; }
            var dlg = new SaveFileDialog { Title = "导出方案", Filter = "TypeRight 方案 (*.json)|*.json", FileName = p.Id + ".json" };
            if (dlg.ShowDialog(this) != true) return;
            _store.Export(p, dlg.FileName);
            _status.Text = "已导出。";
        }

        private void OnRename(object sender, RoutedEventArgs e)
        {
            var p = Selected();
            if (p == null) { MessageBox.Show(this, "先在列表中选中一个方案。"); return; }
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
            if (p == null) { MessageBox.Show(this, "先在列表中选中一个方案。"); return; }
            if (MessageBox.Show(this, "删除方案「" + p.Name + "」？", "确认", MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK) return;
            _store.Delete(p.Id);
            Reload();
        }
    }

    /// <summary>简单输入对话框。</summary>
    public sealed class InputDialog : Window
    {
        private readonly TextBox _box;

        public string Value { get { return _box.Text; } }

        public InputDialog(string title, string initial)
        {
            Title = title;
            Width = 340; SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;

            var panel = new StackPanel { Margin = new Thickness(16) };
            _box = new TextBox { Text = initial ?? "", Padding = new Thickness(4, 3, 4, 3) };
            panel.Children.Add(_box);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
            var ok = new Button { Content = "确定", Padding = new Thickness(14, 3, 14, 3), IsDefault = true };
            ok.Click += (s, e) => { DialogResult = true; };
            var cancel = new Button { Content = "取消", Padding = new Thickness(12, 3, 12, 3), IsCancel = true, Margin = new Thickness(8, 0, 0, 0) };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            panel.Children.Add(buttons);
            Content = panel;
            Loaded += (s, e) => { _box.SelectAll(); _box.Focus(); };
        }
    }
}
