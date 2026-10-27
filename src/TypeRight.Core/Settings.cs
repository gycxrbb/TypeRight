using System;
using System.IO;
using Newtonsoft.Json;

namespace TypeRight
{
    /// <summary>设置（PRD v2 P1-2）：仅三项 + 方案目录。</summary>
    public class AppSettings
    {
        /// <summary>中西文联动（P0-3）：设置中文字体时自动把西文/数字写为默认西文字体。</summary>
        public bool LatinLinkageEnabled { get; set; } = true;

        /// <summary>默认西文字体。</summary>
        public string DefaultLatinFont { get; set; } = "Times New Roman";

        /// <summary>同类选择默认严格模式（false = 默认宽松全忽略？不——默认仍严格，宽松仅勾选生效）。</summary>
        public bool StrictByDefault { get; set; } = true;

        /// <summary>方案存储目录。</summary>
        public string ProfilesDirectory { get; set; } = Profiles.ProfileStoreV2.DefaultDirectory;
    }

    /// <summary>设置读写。损坏时静默重建默认值。</summary>
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
