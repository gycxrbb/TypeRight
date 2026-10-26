using System;
using System.IO;
using Newtonsoft.Json;

namespace TypeRight.Llm
{
    /// <summary>LLM 设置（PRD §4.3）：BaseURL / API Key / 模型名，存 %APPDATA%\TypeRight\settings.json。key 不回传任何服务器（仅发往用户配置的 BaseURL）。</summary>
    public class LlmConfig
    {
        public bool Enabled { get; set; } = false;
        public string BaseUrl { get; set; } = "https://api.openai.com/v1";
        public string ApiKey { get; set; } = "";
        public string Model { get; set; } = "gpt-4o-mini";
        public int TimeoutSeconds { get; set; } = 30;
    }

    public class AppSettings
    {
        public LlmConfig Llm { get; set; } = new LlmConfig();

        /// <summary>上次使用的方案 id。</summary>
        public string LastProfileId { get; set; } = "academic-lecture";

        /// <summary>“同时挂载 Word 样式”默认勾选状态。</summary>
        public bool MountStylesDefault { get; set; } = true;
    }

    /// <summary>设置读写。损坏时静默重建默认值（设置损坏不应阻塞排版主流程）。</summary>
    public static class SettingsStore
    {
        public static readonly string Directory =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TypeRight");

        public static readonly string FilePath = Path.Combine(Directory, "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new AppSettings();
                return JsonConvert.DeserializeObject<AppSettings>(File.ReadAllText(FilePath)) ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }

        public static void Save(AppSettings settings)
        {
            if (settings == null) throw new ArgumentNullException("settings");
            System.IO.Directory.CreateDirectory(Directory);
            File.WriteAllText(FilePath, JsonConvert.SerializeObject(settings, Newtonsoft.Json.Formatting.Indented));
        }
    }
}
