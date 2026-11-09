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

        /// <summary>跳转强调色（十六进制 RRGGBB）。柔和低饱和系，与命中黄色基底协调。</summary>
        public string EmphasisColor { get; set; } = "#E8B34B";

        /// <summary>字体颜色取色器的自定义色板（COLORREF 整数，与系统颜色对话框原样存取）。</summary>
        public int[] CustomFontColors { get; set; } = new int[0];

        /// <summary>强调色 → Word WdColor 的 BGR 整数（R + G*256 + B*65536）。解析失败回退默认蜜橙。</summary>
        public int EmphasisColorRgb
        {
            get
            {
                try
                {
                    string h = (EmphasisColor ?? "").TrimStart('#');
                    if (h.Length != 6) throw new FormatException();
                    int r = Convert.ToInt32(h.Substring(0, 2), 16);
                    int g = Convert.ToInt32(h.Substring(2, 2), 16);
                    int b = Convert.ToInt32(h.Substring(4, 2), 16);
                    return r | (g << 8) | (b << 16);
                }
                catch { return 0xE8 | (0xB3 << 8) | (0x4B << 16); }
            }
        }
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
