using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TypeRight.Profiles
{
    /// <summary>方案 JSON Schema v2（PRD v2 §8）：签名规则制，由用户操作沉淀生成。</summary>
    public class ProfileV2
    {
        [JsonProperty("schema_version")]
        public int SchemaVersion { get; set; } = 2;

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("created")]
        public string Created { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");

        [JsonProperty("rules")]
        public List<ProfileRule> Rules { get; set; } = new List<ProfileRule>();

        [JsonProperty("convertFormulas")]
        public bool ConvertFormulas { get; set; } = true;

        public override string ToString() { return Name ?? Id ?? "(未命名方案)"; }
    }

    public class ProfileRule
    {
        [JsonProperty("order")]
        public int Order { get; set; }

        [JsonProperty("match")]
        public MatchSpec Match { get; set; }

        [JsonProperty("apply")]
        public ApplySpec Apply { get; set; }
    }

    /// <summary>匹配条件：签名值 + 忽略维度；catchAll = 命中所有未被前面规则匹配的段落。</summary>
    public class MatchSpec
    {
        [JsonProperty("eastAsiaFont")]
        public string EastAsiaFont { get; set; }

        [JsonProperty("latinFont")]
        public string LatinFont { get; set; }

        [JsonProperty("sizePt", NullValueHandling = NullValueHandling.Ignore)]
        public double? SizePt { get; set; }

        [JsonProperty("bold", NullValueHandling = NullValueHandling.Ignore)]
        public bool? Bold { get; set; }

        [JsonProperty("color")]
        public string Color { get; set; }          // "#RRGGBB"；null/省略 = 忽略该维度

        [JsonProperty("ignore", NullValueHandling = NullValueHandling.Ignore)]
        public List<string> Ignore { get; set; }   // sizePt / color / bold

        [JsonProperty("catchAll", NullValueHandling = NullValueHandling.Ignore)]
        public bool? CatchAll { get; set; }

        public Signature.FormatSignature.IgnoreDims IgnoreValue
        {
            get
            {
                var dims = Signature.FormatSignature.IgnoreDims.None;
                if (Ignore == null) return dims;
                foreach (string s in Ignore)
                {
                    if (string.Equals(s, "sizePt", StringComparison.OrdinalIgnoreCase)) dims |= Signature.FormatSignature.IgnoreDims.SizePt;
                    else if (string.Equals(s, "color", StringComparison.OrdinalIgnoreCase)) dims |= Signature.FormatSignature.IgnoreDims.Color;
                    else if (string.Equals(s, "bold", StringComparison.OrdinalIgnoreCase)) dims |= Signature.FormatSignature.IgnoreDims.Bold;
                }
                return dims;
            }
        }

        /// <summary>方案规则对段落的匹配：null/省略的维度即忽略（Schema v2 语义）。</summary>
        public bool MatchesPara(Signature.FormatSignature s)
        {
            if (CatchAll == true) return true;
            var ig = IgnoreValue;
            if (EastAsiaFont != null && !string.Equals(s.EastAsiaFont, EastAsiaFont, StringComparison.OrdinalIgnoreCase)) return false;
            if (LatinFont != null && !string.Equals(s.LatinFont, LatinFont, StringComparison.OrdinalIgnoreCase)) return false;
            if (SizePt.HasValue && (ig & Signature.FormatSignature.IgnoreDims.SizePt) == 0)
            {
                if (double.IsNaN(s.SizePt) || Math.Abs(s.SizePt - SizePt.Value) >= 0.01) return false;
            }
            if (Bold.HasValue && (ig & Signature.FormatSignature.IgnoreDims.Bold) == 0 && s.Bold != Bold.Value) return false;
            if (Color != null && (ig & Signature.FormatSignature.IgnoreDims.Color) == 0)
            {
                var hex = System.Text.RegularExpressions.Regex.Match(Color, "^#?([0-9A-Fa-f]{6})$");
                if (!hex.Success || s.ColorRgb != Convert.ToInt32(hex.Groups[1].Value, 16)) return false;
            }
            return true;
        }
    }

    /// <summary>应用格式。null = 保持源值不触碰（与浮条「—」一致）。</summary>
    public class ApplySpec
    {
        [JsonProperty("eastAsiaFont")]
        public string EastAsiaFont { get; set; }

        [JsonProperty("latinFont")]
        public string LatinFont { get; set; }

        [JsonProperty("sizePt", NullValueHandling = NullValueHandling.Include)]
        public double? SizePt { get; set; }

        [JsonProperty("bold", NullValueHandling = NullValueHandling.Include)]
        public bool? Bold { get; set; }

        [JsonProperty("color")]
        public string Color { get; set; }
    }

    /// <summary>方案校验（PRD §8：字段级错误定位）。</summary>
    public static class ProfileV2Validator
    {
        private static readonly Regex HexColor = new Regex("^#[0-9A-Fa-f]{6}$", RegexOptions.Compiled);

        public static List<string> Validate(ProfileV2 profile)
        {
            var errors = new List<string>();
            if (profile == null) { errors.Add("root: 方案内容为空。"); return errors; }
            if (profile.SchemaVersion != 2)
                errors.Add("schema_version: 必须为 2，当前 " + profile.SchemaVersion + "。");
            if (string.IsNullOrWhiteSpace(profile.Id))
                errors.Add("id: 不能为空。");
            if (string.IsNullOrWhiteSpace(profile.Name))
                errors.Add("name: 不能为空。");
            if (profile.Rules == null || profile.Rules.Count == 0)
            {
                errors.Add("rules: 至少需要一条规则。");
                return errors;
            }

            for (int i = 0; i < profile.Rules.Count; i++)
            {
                string at = "rules[" + i + "]";
                ProfileRule rule = profile.Rules[i];
                if (rule == null) { errors.Add(at + ": 规则为空。"); continue; }
                if (rule.Order != i + 1)
                    errors.Add(at + ".order: 必须从 1 起连续编号（当前 " + rule.Order + "，应为 " + (i + 1) + "）。");
                if (rule.Match == null) { errors.Add(at + ".match: 缺失。"); continue; }
                if (rule.Apply == null) { errors.Add(at + ".apply: 缺失。"); continue; }

                bool catchAll = rule.Match.CatchAll == true;
                if (!catchAll)
                {
                    if (string.IsNullOrEmpty(rule.Match.EastAsiaFont) && string.IsNullOrEmpty(rule.Match.LatinFont) &&
                        !rule.Match.SizePt.HasValue && !rule.Match.Bold.HasValue && rule.Match.Color == null)
                        errors.Add(at + ".match: 非兜底规则至少要有一个签名维度。");
                    if (rule.Match.SizePt.HasValue && (rule.Match.SizePt.Value <= 0 || rule.Match.SizePt.Value > 400))
                        errors.Add(at + ".match.sizePt: 超出范围 (0, 400]。");
                    if (rule.Match.Color != null && !HexColor.IsMatch(rule.Match.Color))
                        errors.Add(at + ".match.color: 必须为 #RRGGBB。");
                }
                if (rule.Apply.EastAsiaFont == null && rule.Apply.LatinFont == null && !rule.Apply.SizePt.HasValue &&
                    !rule.Apply.Bold.HasValue && rule.Apply.Color == null)
                    errors.Add(at + ".apply: 全部为 null，规则不会改动任何内容。");
                if (rule.Apply.SizePt.HasValue && (rule.Apply.SizePt.Value <= 0 || rule.Apply.SizePt.Value > 400))
                    errors.Add(at + ".apply.sizePt: 超出范围 (0, 400]。");
                if (rule.Apply.Color != null && !HexColor.IsMatch(rule.Apply.Color))
                    errors.Add(at + ".apply.color: 必须为 #RRGGBB。");
                if (catchAll && i != profile.Rules.Count - 1)
                    errors.Add(at + ".match.catchAll: 兜底规则必须位于末位。");
                if (rule.Match.Ignore != null)
                {
                    var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "sizePt", "color", "bold" };
                    foreach (string ig in rule.Match.Ignore)
                        if (!known.Contains(ig))
                            errors.Add(at + ".match.ignore: 未知维度 \"" + ig + "\"（可用 sizePt/color/bold）。");
                }
            }
            return errors;
        }
    }

    /// <summary>方案读写/导入导出（PRD P1-1）。目录默认 %APPDATA%\TypeRight\profiles\。</summary>
    public sealed class ProfileStoreV2
    {
        public static readonly string DefaultDirectory =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TypeRight", "profiles");

        public string Directory { get; private set; }

        public ProfileStoreV2(string directory = null)
        {
            Directory = string.IsNullOrEmpty(directory) ? DefaultDirectory : directory;
            System.IO.Directory.CreateDirectory(Directory);
        }

        /// <summary>列出全部方案；非法文件被拒绝并收集字段级错误。</summary>
        public List<ProfileV2> LoadAll(List<string> errors)
        {
            var result = new List<ProfileV2>();
            foreach (string file in System.IO.Directory.GetFiles(Directory, "*.json")
                         .OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
            {
                ProfileV2 p;
                List<string> fileErrors;
                if (TryLoad(file, out p, out fileErrors)) result.Add(p);
                else errors.Add(Path.GetFileName(file) + " → " + string.Join("；", fileErrors));
            }
            return result;
        }

        public static bool TryLoad(string file, out ProfileV2 profile, out List<string> errors)
        {
            profile = null;
            errors = new List<string>();
            try
            {
                string json = File.ReadAllText(file, Encoding.UTF8);
                profile = JsonConvert.DeserializeObject<ProfileV2>(json);
                errors = ProfileV2Validator.Validate(profile);
                if (errors.Count > 0) { profile = null; return false; }
                return true;
            }
            catch (Exception ex)
            {
                errors.Add("JSON 解析失败: " + ex.Message);
                return false;
            }
        }

        public string Save(ProfileV2 profile)
        {
            var errors = ProfileV2Validator.Validate(profile);
            if (errors.Count > 0)
                throw new ProfileValidationException("方案校验未通过，已拒绝保存。", errors);

            string file = Path.Combine(Directory, profile.Id + ".json");
            if (File.Exists(file)) File.Copy(file, file + ".bak", true);
            File.WriteAllText(file, JsonConvert.SerializeObject(profile, Newtonsoft.Json.Formatting.Indented),
                new UTF8Encoding(false));
            return file;
        }

        public bool Delete(string id)
        {
            string file = Path.Combine(Directory, id + ".json");
            if (File.Exists(file)) { File.Delete(file); return true; }
            return false;
        }

        public void Rename(ProfileV2 profile, string newName)
        {
            profile.Name = newName;
            Save(profile);
        }

        public void Export(ProfileV2 profile, string targetFile)
        {
            File.WriteAllText(targetFile, JsonConvert.SerializeObject(profile, Newtonsoft.Json.Formatting.Indented),
                new UTF8Encoding(false));
        }

        public string Import(string sourceFile, out List<string> errors)
        {
            ProfileV2 p;
            if (!TryLoad(sourceFile, out p, out errors)) return null;
            // 同 id 去重：追加 -imported
            if (File.Exists(Path.Combine(Directory, p.Id + ".json")))
            {
                int n = 2;
                while (File.Exists(Path.Combine(Directory, p.Id + "-" + n + ".json"))) n++;
                p.Id = p.Id + "-" + n;
            }
            Save(p);
            return p.Id;
        }
    }

    [Serializable]
    public sealed class ProfileValidationException : Exception
    {
        public List<string> Errors { get; private set; }

        public ProfileValidationException(string message, List<string> errors)
            : base(message)
        {
            Errors = errors ?? new List<string>();
        }
    }
}
