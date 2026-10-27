using System;
using System.Windows;

namespace TypeRight.Addin.UI
{
    public partial class SettingsWindow : Window
    {
        private readonly AppSettings _settings;

        public SettingsWindow(AppSettings settings)
        {
            InitializeComponent();
            _settings = settings ?? new AppSettings();
            ChkLinkage.IsChecked = _settings.LatinLinkageEnabled;
            TxtLatin.Text = _settings.DefaultLatinFont;
            ChkStrict.IsChecked = _settings.StrictByDefault;
            TxtProfilesDir.Text = _settings.ProfilesDirectory;
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            _settings.LatinLinkageEnabled = ChkLinkage.IsChecked == true;
            _settings.DefaultLatinFont = string.IsNullOrWhiteSpace(TxtLatin.Text) ? "Times New Roman" : TxtLatin.Text.Trim();
            _settings.StrictByDefault = ChkStrict.IsChecked == true;
            _settings.ProfilesDirectory = string.IsNullOrWhiteSpace(TxtProfilesDir.Text)
                ? Profiles.ProfileStoreV2.DefaultDirectory
                : TxtProfilesDir.Text.Trim();
            SettingsStore.Save(_settings);
            Close();
        }
    }
}
