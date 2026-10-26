using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace TypeRight.Profiles
{
    /// <summary>对齐方式。none 表示方案不管对齐（保留原样）。</summary>
    public enum ProfileAlignment { None, Left, Center, Right, Justify }

    /// <summary>全半角标点统一策略（PRD §3 mixingRules.unifyPunctuation）。</summary>
    public enum PunctuationMode { None, Fullwidth, Halfwidth }

    /// <summary>排版方案（PRD §3 JSON Schema）。</summary>
    public class Profile
    {
        [JsonProperty("schema_version")]
        public int SchemaVersion { get; set; } = 1;

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("roles")]
        public Dictionary<string, RoleFormat> Roles { get; set; } = new Dictionary<string, RoleFormat>();

        [JsonProperty("formula")]
        public FormulaOptions Formula { get; set; } = new FormulaOptions();

        [JsonProperty("mixingRules")]
        public MixingRuleOptions MixingRules { get; set; } = new MixingRuleOptions();

        public override string ToString() { return Name ?? Id ?? "(unnamed profile)"; }
    }

    /// <summary>单个角色的格式定义。字体只写对应通道：eastAsiaFont→NameFarEast，latinFont→NameAscii/NameOther。</summary>
    public class RoleFormat
    {
        [JsonProperty("label")]
        public string Label { get; set; }

        [JsonProperty("eastAsiaFont")]
        public string EastAsiaFont { get; set; }

        [JsonProperty("latinFont")]
        public string LatinFont { get; set; }

        [JsonProperty("sizePt", NullValueHandling = NullValueHandling.Ignore)]
        public double? SizePt { get; set; }

        [JsonProperty("bold", NullValueHandling = NullValueHandling.Ignore)]
        public bool? Bold { get; set; }

        [JsonProperty("color")]
        public string Color { get; set; }

        [JsonProperty("alignment")]
        public string Alignment { get; set; }

        [JsonProperty("wordStyle")]
        public string WordStyle { get; set; }

        [JsonProperty("outlineLevel", NullValueHandling = NullValueHandling.Ignore)]
        public int? OutlineLevel { get; set; }

        [JsonProperty("lineSpacing", NullValueHandling = NullValueHandling.Ignore)]
        public double? LineSpacing { get; set; }

        [JsonProperty("spaceBeforePt", NullValueHandling = NullValueHandling.Ignore)]
        public double? SpaceBeforePt { get; set; }

        [JsonProperty("spaceAfterPt", NullValueHandling = NullValueHandling.Ignore)]
        public double? SpaceAfterPt { get; set; }

        [JsonProperty("firstLineIndentPt", NullValueHandling = NullValueHandling.Ignore)]
        public double? FirstLineIndentPt { get; set; }

        [JsonIgnore]
        public ProfileAlignment AlignmentValue
        {
            get
            {
                if (string.IsNullOrEmpty(Alignment)) return ProfileAlignment.None;
                ProfileAlignment a;
                return Enum.TryParse(Alignment, true, out a) ? a : ProfileAlignment.None;
            }
        }

        // ---- DataGrid 编辑用的字符串视图属性（JSON 序列化忽略） ----

        [JsonIgnore]
        public string SizePtText
        {
            get { return SizePt.HasValue ? SizePt.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) : ""; }
            set { SizePt = ParseDouble(value); }
        }

        [JsonIgnore]
        public string BoldText
        {
            get { return !Bold.HasValue ? "" : (Bold.Value ? "1" : "0"); }
            set
            {
                string v = (value ?? "").Trim();
                Bold = v.Length == 0 ? (bool?)null : (v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase));
            }
        }

        [JsonIgnore]
        public string LineSpacingText
        {
            get { return LineSpacing.HasValue ? LineSpacing.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) : ""; }
            set { LineSpacing = ParseDouble(value); }
        }

        [JsonIgnore]
        public string FirstLineIndentPtText
        {
            get { return FirstLineIndentPt.HasValue ? FirstLineIndentPt.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) : ""; }
            set { FirstLineIndentPt = ParseDouble(value); }
        }

        private static double? ParseDouble(string s)
        {
            double d;
            if (double.TryParse((s ?? "").Trim(), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out d)) return d;
            return null;
        }
    }

    public class FormulaOptions
    {
        [JsonProperty("convertPlainText")]
        public bool ConvertPlainText { get; set; } = true;

        [JsonProperty("mathFont")]
        public string MathFont { get; set; } = "Cambria Math";

        [JsonProperty("sizePt", NullValueHandling = NullValueHandling.Ignore)]
        public double? SizePt { get; set; }
    }

    public class MixingRuleOptions
    {
        [JsonProperty("autoSpaceCjkLatin")]
        public bool AutoSpaceCjkLatin { get; set; } = false;

        [JsonProperty("unifyPunctuation")]
        public string UnifyPunctuation { get; set; } = "none";

        [JsonProperty("digitsToLatinFont")]
        public bool DigitsToLatinFont { get; set; } = true;

        [JsonIgnore]
        public PunctuationMode PunctuationMode
        {
            get
            {
                PunctuationMode m;
                return Enum.TryParse(UnifyPunctuation, true, out m) ? m : PunctuationMode.None;
            }
        }
    }

    /// <summary>方案校验：缺字段/非法值时拒绝加载（PRD §3 规则），不 fallback 到半成品。</summary>
    public static class ProfileValidator
    {
        private static readonly Regex HexColor = new Regex("^#[0-9A-Fa-f]{6}$", RegexOptions.Compiled);
        // 字体名：非空、无路径字符、无控制字符
        private static readonly Regex FontNamePattern = new Regex("^[^\\\\/:*?\"<>|\r\n\x00-\x1F]{1,63}$", RegexOptions.Compiled);

        private static readonly HashSet<string> KnownRoleKeys = new HashSet<string>
        {
            "title", "heading1", "heading2", "heading3", "body", "caption",
            "listItem", "tableText", "formulaLine"
        };

        public static List<string> Validate(Profile profile)
        {
            var errors = new List<string>();
            if (profile == null)
            {
                errors.Add("方案内容为空。");
                return errors;
            }
            if (profile.SchemaVersion != 1)
                errors.Add("schema_version 必须为 1，当前为 " + profile.SchemaVersion + "。");
            if (string.IsNullOrWhiteSpace(profile.Id))
                errors.Add("缺少 id 字段。");
            if (string.IsNullOrWhiteSpace(profile.Name))
                errors.Add("缺少 name 字段。");
            if (profile.Roles == null || profile.Roles.Count == 0)
            {
                errors.Add("缺少 roles 字段或 roles 为空。");
                return errors;
            }
            if (!profile.Roles.ContainsKey("body"))
                errors.Add("roles 必须包含 body（正文）角色。");

            foreach (var kv in profile.Roles)
            {
                string key = kv.Key;
                if (!KnownRoleKeys.Contains(key))
                {
                    errors.Add(string.Format("roles 含未知角色键 {0}。", key));
                    continue;
                }
                RoleFormat rf = kv.Value;
                if (rf == null)
                {
                    errors.Add(string.Format("roles.{0} 内容为空。", key));
                    continue;
                }
                if (!string.IsNullOrEmpty(rf.EastAsiaFont) && !FontNamePattern.IsMatch(rf.EastAsiaFont))
                    errors.Add(string.Format("roles.{0}.eastAsiaFont 字体名非法。", key));
                if (!string.IsNullOrEmpty(rf.LatinFont) && !FontNamePattern.IsMatch(rf.LatinFont))
                    errors.Add(string.Format("roles.{0}.latinFont 字体名非法。", key));
                if (rf.SizePt.HasValue && (rf.SizePt.Value <= 0 || rf.SizePt.Value > 400))
                    errors.Add(string.Format("roles.{0}.sizePt 超出范围 (0, 400]。", key));
                if (!string.IsNullOrEmpty(rf.Color) && !HexColor.IsMatch(rf.Color))
                    errors.Add(string.Format("roles.{0}.color 必须为 #RRGGBB 格式。", key));
                if (rf.OutlineLevel.HasValue && (rf.OutlineLevel.Value < 1 || rf.OutlineLevel.Value > 9))
                    errors.Add(string.Format("roles.{0}.outlineLevel 必须在 1..9。", key));
                if (rf.LineSpacing.HasValue && (rf.LineSpacing.Value <= 0 || rf.LineSpacing.Value > 10))
                    errors.Add(string.Format("roles.{0}.lineSpacing 超出范围 (0, 10]。", key));
                if (!string.IsNullOrEmpty(rf.Alignment))
                {
                    ProfileAlignment a;
                    bool ok = Enum.TryParse(rf.Alignment, true, out a);
                    if (!ok || (a == ProfileAlignment.None && !string.Equals(rf.Alignment, "none", StringComparison.OrdinalIgnoreCase)))
                        errors.Add(string.Format("roles.{0}.alignment 取值非法：{1}。", key, rf.Alignment));
                }
            }
            return errors;
        }

        /// <summary>字体未安装仅告警（Word 会静默替换），语法非法才拒绝。</summary>
        public static List<string> WarnNotInstalledFonts(Profile profile)
        {
            var warnings = new List<string>();
            var installed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (System.Drawing.Text.InstalledFontCollection fonts = new System.Drawing.Text.InstalledFontCollection())
                {
                    foreach (System.Drawing.FontFamily f in fonts.Families) installed.Add(f.Name);
                }
            }
            catch { /* 读取系统字体失败时跳过告警 */ }
            if (installed.Count == 0) return warnings;
            if (profile.Roles == null) return warnings;
            foreach (var kv in profile.Roles)
            {
                RoleFormat rf = kv.Value;
                if (rf == null) continue;
                CheckFont(rf.EastAsiaFont, kv.Key + ".eastAsiaFont", installed, warnings);
                CheckFont(rf.LatinFont, kv.Key + ".latinFont", installed, warnings);
            }
            return warnings;
        }

        private static void CheckFont(string font, string where, HashSet<string> installed, List<string> warnings)
        {
            if (string.IsNullOrEmpty(font)) return;
            if (!installed.Contains(font) && !installed.Contains(font.Replace(" ", "")))
                warnings.Add(string.Format("字体 {0}（{1}）未在系统字体列表中找到，Word 将使用替代字体。", font, where));
        }
    }
}
