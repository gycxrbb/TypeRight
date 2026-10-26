using System;
using System.Windows;
using TypeRight.Llm;

namespace TypeRight.Addin.Windows
{
    public partial class SettingsWindow : Window
    {
        private readonly AppSettings _settings;

        public SettingsWindow(AppSettings settings)
        {
            InitializeComponent();
            _settings = settings ?? new AppSettings();
            ChkEnabled.IsChecked = _settings.Llm.Enabled;
            TxtBaseUrl.Text = _settings.Llm.BaseUrl;
            TxtModel.Text = _settings.Llm.Model;
            TxtTimeout.Text = _settings.Llm.TimeoutSeconds.ToString();
            ChkMountStyles.IsChecked = _settings.MountStylesDefault;
            // PasswordBox 不支持绑定，从设置恢复
            if (!string.IsNullOrEmpty(_settings.Llm.ApiKey))
                TxtApiKey.Password = _settings.Llm.ApiKey;
        }

        private void OnTest(object sender, RoutedEventArgs e)
        {
            TxtTestResult.Text = "";
            var config = Collect();
            if (string.IsNullOrWhiteSpace(config.ApiKey))
            {
                TxtTestResult.Text = "未填写 API Key";
                TxtTestResult.Foreground = System.Windows.Media.Brushes.Firebrick;
                return;
            }
            TxtTestResult.Foreground = System.Windows.Media.Brushes.Gray;
            TxtTestResult.Text = "请求中…";
            try
            {
                var client = new LlmClient(Collect());
                string role = client.Ping();
                TxtTestResult.Foreground = System.Windows.Media.Brushes.SeaGreen;
                TxtTestResult.Text = "连通正常，示例判定：" + role;
            }
            catch (Exception ex)
            {
                TxtTestResult.Foreground = System.Windows.Media.Brushes.Firebrick;
                TxtTestResult.Text = "失败：" + ex.Message;
            }
        }

        private LlmConfig Collect()
        {
            int timeout;
            if (!int.TryParse(TxtTimeout.Text.Trim(), out timeout) || timeout <= 0) timeout = 30;
            return new LlmConfig
            {
                Enabled = ChkEnabled.IsChecked == true,
                BaseUrl = TxtBaseUrl.Text.Trim(),
                ApiKey = TxtApiKey.Password,
                Model = TxtModel.Text.Trim(),
                TimeoutSeconds = timeout
            };
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            _settings.Llm = Collect();
            _settings.MountStylesDefault = ChkMountStyles.IsChecked == true;
            SettingsStore.Save(_settings);
            Close();
        }
    }
}
