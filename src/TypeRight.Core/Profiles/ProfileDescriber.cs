using System;
using System.Collections.Generic;
using System.Linq;
using TypeRight.Formatting;

namespace TypeRight.Profiles
{
    /// <summary>一条规则的人话描述（供方案卡片直接渲染，不含任何 UI 依赖）。
    ///
    /// 注意：这里必须是 **属性** 而不是字段——WPF 的 {Binding} 只认 CLR 属性，
    /// 绑到 public field 上会静默失败（不报错、不抛异常，界面就是一片空白）。
    /// 2026-11 曾因此让方案卡片只剩按钮、文字全空。</summary>
    public sealed class RuleSummary
    {
        /// <summary>挑中什么样的段落，例如「仿宋 · 黑色 · 四号 · 不加粗」。</summary>
        public string MatchText { get; set; }

        /// <summary>改了什么，例如「字号 → 三号」。这是用户真正关心的那一句。</summary>
        public string ChangeText { get; set; }

        /// <summary>是否兜底规则（其余段落）。</summary>
        public bool IsCatchAll { get; set; }

        /// <summary>匹配条件是否几乎不挑（无字体/字号/加粗/颜色任何一个维度）。</summary>
        public bool IsBroad { get; set; }

        /// <summary>一句话总述，例如「仿宋 · 黑色 · 四号 · 不加粗 → 字号改成三号」。</summary>
        public string Sentence
        {
            get
            {
                if (string.IsNullOrEmpty(ChangeText)) return MatchText;
                return MatchText + "\n" + ChangeText;
            }
        }
    }

    /// <summary>整个方案的人话描述。</summary>
    public sealed class ProfileSummary
    {
        public string Name { get; set; }

        /// <summary>「改了 1 处」这类计数文案（按"变化点"算，不按规则条数）。</summary>
        public string RuleCountText { get; set; }

        public List<RuleSummary> Rules { get; set; }

        /// <summary>「会同时把全文公式转成公式对象」之类；无则 null。</summary>
        public string FormulaNote { get; set; }

        public ProfileSummary()
        {
            Rules = new List<RuleSummary>();
        }
    }

    /// <summary>
    /// 方案 → 人话。
    ///
    /// 产品立意（用户原话）：方案记录的是「我把某一类文本的某个格式属性从 X 改成了 Y」。
    /// 所以卡片必须优先回答 **改了什么**（字号 四号 → 三号），而不是罗列完整格式栈。
    /// 完整格式栈只是"这些段落长什么样"的识别依据，属于次要信息。
    ///
    /// 纯逻辑、无 Word 依赖 → xUnit 直测。字号一律走 ChineseFontSizes，
    /// 与浮条字号框同一口径（16 磅显示「三号」而不是裸数字）。
    /// </summary>
    public static class ProfileDescriber
    {
        // ───────────────────────── 整方案 ─────────────────────────

        public static ProfileSummary Describe(ProfileV2 profile)
        {
            var summary = new ProfileSummary();
            if (profile == null) return summary;

            summary.Name = string.IsNullOrWhiteSpace(profile.Name) ? "(未命名方案)" : profile.Name;

            var rules = (profile.Rules ?? new List<ProfileRule>())
                .Where(r => r != null)
                .OrderBy(r => r.Order)
                .ToList();

            foreach (ProfileRule rule in rules)
                summary.Rules.Add(DescribeRule(rule));

            int changeCount = summary.Rules.Count(r => !string.IsNullOrEmpty(r.ChangeText));
            summary.RuleCountText = changeCount <= 1
                ? "改了 1 处格式"
                : "改了 " + changeCount + " 处格式";

            if (profile.ConvertFormulas) summary.FormulaNote = "顺带把全文公式转成公式对象";
            return summary;
        }

        // ───────────────────────── 单条规则 ─────────────────────────

        public static RuleSummary DescribeRule(ProfileRule rule)
        {
            var rs = new RuleSummary();
            if (rule == null) return rs;

            bool catchAll = rule.Match != null && rule.Match.CatchAll == true;
            rs.IsCatchAll = catchAll;

            if (catchAll)
            {
                rs.MatchText = "其余所有段落";
            }
            else
            {
                rs.MatchText = DescribeMatch(rule.Match);
                rs.IsBroad = IsBroad(rule.Match);
            }

            rs.ChangeText = DescribeChange(rule.Match, rule.Apply);
            return rs;
        }

        /// <summary>
        /// 「改了什么」——这是整个卡片最重要的一行。
        /// 逐一比对 match（原格式）与 apply（目标格式），只讲真正发生变化的属性，
        /// 并带上「从什么变成什么」，用户才认得出这是自己做过的那次操作。
        /// </summary>
        public static string DescribeChange(MatchSpec match, ApplySpec apply)
        {
            if (apply == null) return null;

            var changes = new List<string>();

            // 字体
            if (apply.EastAsiaFont != null)
            {
                string from = match != null ? FirstNonNull(match.EastAsiaFont, match.LatinFont) : null;
                changes.Add(SameText(from, apply.EastAsiaFont)
                    ? "中文字体 " + apply.EastAsiaFont
                    : "中文字体 " + (from ?? "?") + " → " + apply.EastAsiaFont);
            }
            // 字号（最常见的场景，单独成句）
            if (apply.SizePt.HasValue)
            {
                string to = SizeLabel(apply.SizePt.Value);
                if (match != null && match.SizePt.HasValue)
                {
                    double fromPt = match.SizePt.Value;
                    bool same = Math.Abs(fromPt - apply.SizePt.Value) < 0.01;
                    changes.Add(same ? "字号 " + to : "字号 " + SizeLabel(fromPt) + " → " + to);
                }
                else
                {
                    changes.Add("字号 → " + to);
                }
            }
            // 加粗
            if (apply.Bold.HasValue)
            {
                string to = apply.Bold.Value ? "加粗" : "取消加粗";
                if (match != null && match.Bold.HasValue && match.Bold.Value == apply.Bold.Value)
                    changes.Add(to);
                else if (match != null && match.Bold.HasValue)
                    changes.Add((match.Bold.Value ? "加粗" : "不加粗") + " → " + to);
                else
                    changes.Add(to);
            }
            // 颜色
            if (apply.Color != null)
            {
                string to = ColorLabel(apply.Color);
                if (match != null && match.Color != null
                    && string.Equals(match.Color, apply.Color, StringComparison.OrdinalIgnoreCase))
                    changes.Add("颜色 " + to);
                else if (match != null && match.Color != null)
                    changes.Add("颜色 " + ColorLabel(match.Color) + " → " + to);
                else
                    changes.Add("颜色 → " + to);
            }
            // 西文字体：只有与中文字体不同才值得单独说
            if (apply.LatinFont != null)
            {
                bool sameAsMatch = match != null && SameText(match.LatinFont, apply.LatinFont);
                if (!sameAsMatch)
                    changes.Add("西文字体 → " + apply.LatinFont);
            }

            return changes.Count == 0 ? null : string.Join("，", changes.ToArray());
        }

        // ───────────────────────── 匹配条件 ─────────────────────────

        /// <summary>匹配条件 → 人话。这是"这些段落长什么样"的识别依据。
        /// 顺序刻意与用户读出的方式一致：字体 → 颜色 → 字号 → 加粗，与用户原话
        /// 「仿宋 黑色 四号 无加粗」同序。</summary>
        public static string DescribeMatch(MatchSpec m)
        {
            if (m == null) return "任意段落";

            var parts = new List<string>();
            string font = FirstNonNull(m.EastAsiaFont, m.LatinFont);
            if (font != null) parts.Add(font);
            if (m.Color != null) parts.Add(ColorLabel(m.Color));
            if (m.SizePt.HasValue) parts.Add(SizeLabel(m.SizePt.Value));
            if (m.Bold.HasValue) parts.Add(m.Bold.Value ? "加粗" : "不加粗");
            if (m.InTable == true) parts.Add("表格内");
            else if (m.InTable == false) parts.Add("正文");

            string text = parts.Count == 0 ? "任意段落" : string.Join(" · ", parts.ToArray());

            // 宽松模式：把被忽略的维度讲出来，否则用户不知道它为什么"看起来更宽"
            string ignores = DescribeIgnore(m.Ignore);
            if (ignores != null) text += "（不看" + ignores + "）";
            return text;
        }

        /// <summary>目标格式全貌（次要信息，用于悬停详述）。</summary>
        public static string DescribeApply(ApplySpec a)
        {
            if (a == null) return "不改变";

            var parts = new List<string>();
            if (a.EastAsiaFont != null) parts.Add(a.EastAsiaFont);
            if (a.SizePt.HasValue) parts.Add(SizeLabel(a.SizePt.Value));
            if (a.Bold.HasValue) parts.Add(a.Bold.Value ? "加粗" : "不加粗");
            if (a.Color != null) parts.Add(ColorLabel(a.Color));
            if (a.LatinFont != null) parts.Add("西文 " + a.LatinFont);

            return parts.Count == 0 ? "不改变" : string.Join(" · ", parts.ToArray());
        }

        public static string DescribeIgnore(List<string> ignore)
        {
            if (ignore == null || ignore.Count == 0) return null;
            var names = new List<string>();
            foreach (string s in ignore)
            {
                if (string.Equals(s, "sizePt", StringComparison.OrdinalIgnoreCase)) AddOnce(names, "字号");
                else if (string.Equals(s, "color", StringComparison.OrdinalIgnoreCase)) AddOnce(names, "颜色");
                else if (string.Equals(s, "bold", StringComparison.OrdinalIgnoreCase)) AddOnce(names, "加粗");
            }
            return names.Count == 0 ? null : string.Join("、", names.ToArray());
        }

        /// <summary>匹配条件是否几乎不挑。</summary>
        public static bool IsBroad(MatchSpec m)
        {
            if (m == null) return true;
            if (m.CatchAll == true) return true;
            return m.EastAsiaFont == null && m.LatinFont == null
                && !m.SizePt.HasValue && !m.Bold.HasValue && m.Color == null;
        }

        // ───────────────────────── 小工具 ─────────────────────────

        /// <summary>字号 → 「三号」。刻意不带「（16 磅）」——卡片上要短，
        /// 磅值放在悬停提示里（见 SizeLabelLong）。</summary>
        public static string SizeLabel(double pt)
        {
            return ChineseFontSizes.NameOf(pt) ?? (ChineseFontSizes.PtText(pt) + " 磅");
        }

        /// <summary>字号 → 「三号（16 磅）」，给悬停详述用。</summary>
        public static string SizeLabelLong(double pt)
        {
            string name = ChineseFontSizes.NameOf(pt);
            return name != null
                ? name + "（" + ChineseFontSizes.PtText(pt) + " 磅）"
                : ChineseFontSizes.PtText(pt) + " 磅";
        }

        public static string ColorLabel(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return "自动";
            switch (hex.Trim().ToUpperInvariant().TrimStart('#'))
            {
                case "000000": return "黑色";
                case "FFFFFF": return "白色";
                case "FF0000": return "红色";
                case "1E1E1E": return "近黑";
                case "333333": return "深灰";
                case "808080": return "灰色";
                default: return hex;
            }
        }

        private static bool SameText(string a, string b)
        {
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        private static void AddOnce(List<string> list, string v)
        {
            if (!list.Contains(v)) list.Add(v);
        }

        private static string FirstNonNull(string a, string b)
        {
            if (!string.IsNullOrEmpty(a)) return a;
            if (!string.IsNullOrEmpty(b)) return b;
            return null;
        }
    }
}
