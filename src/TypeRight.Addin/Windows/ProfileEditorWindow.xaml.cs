using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using TypeRight.Addin.Controllers;
using TypeRight.Profiles;

namespace TypeRight.Addin.Windows
{
    public partial class ProfileEditorWindow : Window
    {
        private readonly FormattingController _controller;
        private readonly ProfileStore _store = new ProfileStore();
        private List<LoadedProfile> _profiles = new List<LoadedProfile>();
        private LoadedProfile _selected;

        public ProfileEditorWindow(FormattingController controller)
        {
            InitializeComponent();
            _controller = controller;
            ReloadList();
        }

        private void ReloadList(string selectId = null)
        {
            _controller.ReloadProfiles();
            _profiles = _controller.Profiles;
            LstProfiles.ItemsSource = _profiles;
            if (_profiles.Count > 0)
            {
                int idx = 0;
                if (selectId != null)
                {
                    for (int i = 0; i < _profiles.Count; i++)
                        if (_profiles[i].Profile.Id == selectId) { idx = i; break; }
                }
                LstProfiles.SelectedIndex = idx;
            }
        }

        private void OnProfileSelected(object sender, SelectionChangedEventArgs e)
        {
            _selected = LstProfiles.SelectedItem as LoadedProfile;
            if (_selected == null) return;
            TxtName.Text = _selected.Profile.Name;
            TxtId.Text = _selected.Profile.Id;
            TxtDescription.Text = _selected.Profile.Description ?? "";
            GridRoles.ItemsSource = _selected.Profile.Roles.ToList();
            TxtStatus.Text = _selected.IsBuiltIn ? "内置方案只读；保存将另存为新方案" : "";
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            if (_selected == null) return;
            var profile = _selected.Profile;
            profile.Name = TxtName.Text.Trim();
            profile.Description = TxtDescription.Text.Trim();

            bool builtin = _selected.IsBuiltIn;
            string newId = TxtId.Text.Trim();
            if (builtin || newId != _selected.Profile.Id)
            {
                if (builtin && newId == profile.Id) newId = profile.Id + "-custom";
                profile.Id = newId;
            }
            if (profile.Name.Length == 0 || profile.Id.Length == 0)
            {
                MessageBox.Show("名称与 id 不能为空。");
                return;
            }

            try
            {
                string file = _store.Save(profile);
                TxtStatus.Text = "已保存：" + file;
                ReloadList(profile.Id);
            }
            catch (ProfileValidationException ex)
            {
                MessageBox.Show(string.Join("\r\n", ex.Errors), "校验未通过");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "保存失败");
            }
        }

        private void OnDuplicate(object sender, RoutedEventArgs e)
        {
            try
            {
                string newId = "academic-lecture-copy";
                int n = 2;
                while (System.IO.File.Exists(System.IO.Path.Combine(ProfileStore.ProfilesDirectory, newId + ".json")))
                    newId = "academic-lecture-copy" + (n++);
                string file = _store.DuplicateBuiltIn(newId, "学术讲义/论文 副本" + (n > 2 ? " " + n : ""));
                TxtStatus.Text = "已创建：" + file;
                ReloadList(newId);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "新建失败");
            }
        }

        private void OnImport(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "导入排版方案 JSON",
                Filter = "TypeRight 方案 (*.json)|*.json"
            };
            if (dlg.ShowDialog(this) != true) return;

            Profile p;
            List<string> errors;
            if (!ProfileStore.TryLoad(dlg.FileName, out p, out errors))
            {
                MessageBox.Show("方案被拒绝加载：\r\n" + string.Join("\r\n", errors), "校验未通过");
                return;
            }
            try
            {
                _store.Save(p);
                TxtStatus.Text = "已导入：" + p.Id;
                ReloadList(p.Id);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "导入失败");
            }
        }

        private void OnExport(object sender, RoutedEventArgs e)
        {
            if (_selected == null) return;
            var dlg = new SaveFileDialog
            {
                Title = "导出排版方案",
                Filter = "TypeRight 方案 (*.json)|*.json",
                FileName = _selected.Profile.Id + ".json"
            };
            if (dlg.ShowDialog(this) != true) return;
            try
            {
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(_selected.Profile, Newtonsoft.Json.Formatting.Indented);
                System.IO.File.WriteAllText(dlg.FileName, json, new System.Text.UTF8Encoding(false));
                TxtStatus.Text = "已导出。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "导出失败");
            }
        }

        private void OnDelete(object sender, RoutedEventArgs e)
        {
            if (_selected == null || _selected.IsBuiltIn)
            {
                MessageBox.Show("内置方案不可删除。");
                return;
            }
            if (MessageBox.Show("删除方案 “" + _selected.Profile.Name + "”？", "确认",
                MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK) return;
            try
            {
                if (_store.Delete(_selected.Profile.Id)) TxtStatus.Text = "已删除。";
                ReloadList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "删除失败");
            }
        }
    }
}
