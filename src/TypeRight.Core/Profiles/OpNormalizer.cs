using System;
using System.Collections.Generic;
using System.Linq;
using TypeRight.Signature;

namespace TypeRight.Profiles
{
    /// <summary>
    /// 操作录制 → 方案规则的归一化（纯逻辑，无 Word 依赖，xUnit 直测）。
    ///
    /// 为什么需要：用户在一次「同类选择」会话里通常连着改好几次（字体→字号→加粗→颜色），
    /// 每次 ApplyChanges 都记一条 op，且都带**会话开始时的原始签名**。直接逐条转规则，
    /// 就会产出若干条 match 完全相同、apply 各改一个属性的规则。
    ///
    /// 回放时占用语义是「每段只被首条命中规则处理」（consumed 集合），于是第一条把段落全吃掉，
    /// 后面几条命中 0 段——用户改的字号/加粗全部丢失，方案回放结果与当初手动做的不一致。
    ///
    /// 正确做法：**同一匹配条件的多次修改合并成一条**，各属性取最后一次写入（last-write-wins），
    /// 这正是用户心中"这一类段落最终应该长什么样"的净效果。
    /// </summary>
    public static class OpNormalizer
    {
        /// <summary>两条 match 的匹配条件是否语义等价（用于判定"同一次会话的连续修改"）。
        /// 注意：比较的是 MatchSpec 本身，不是 FormatSignature —— 方案层只认 MatchSpec。</summary>
        public static bool SameMatch(MatchSpec a, MatchSpec b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null) return false;

            // 兜底规则独立成条，绝不与普通规则合并
            if (a.CatchAll == true || b.CatchAll == true) return false;
            if (a.CatchAll != b.CatchAll) return false;

            // 文本域（正文/表格）是分区条件：不同域绝不合并
            if (a.InTable != b.InTable) return false;

            if (!EqIgnoreCase(a.EastAsiaFont, b.EastAsiaFont)) return false;
            if (!EqIgnoreCase(a.LatinFont, b.LatinFont)) return false;
            if (!EqNullable(a.SizePt, b.SizePt)) return false;
            if (!EqNullable(a.Bold, b.Bold)) return false;
            if (!EqIgnoreCase(a.Color, b.Color)) return false;

            // 忽略维度不同 = 匹配口径不同，不能合并
            if (!SameIgnore(a.Ignore, b.Ignore)) return false;
            return true;
        }

        /// <summary>
        /// 把「签名 + 应用格式 + 是否兜底」的原始操作序列归一化成规则列表。
        /// 相邻且匹配条件相同的操作合并；不同匹配条件保持顺序。
        ///
        /// 为什么只合并**相邻**的：中间若夹了别的规则，说明用户切换过选择对象，
        /// 合并会打乱语义顺序（后面那条可能依赖前面已消费的段落）。
        /// </summary>
        public static List<ProfileRule> Normalize(IList<ProfileRule> rules)
        {
            var result = new List<ProfileRule>();
            if (rules == null) return result;

            foreach (ProfileRule rule in rules)
            {
                if (rule == null) continue;
                ProfileRule last = result.Count > 0 ? result[result.Count - 1] : null;

                if (last != null && SameMatch(last.Match, rule.Match) && last.Match.CatchAll != true)
                {
                    // 合并进上一条：apply 各字段独立 last-write-wins（null = 该次没动这个属性）
                    last.Apply = MergeApply(last.Apply, rule.Apply);
                }
                else
                {
                    result.Add(new ProfileRule
                    {
                        Match = rule.Match,
                        Apply = rule.Apply
                    });
                }
            }

            // 重新编号，保证 order 从 1 起连续（校验器要求）
            for (int i = 0; i < result.Count; i++) result[i].Order = i + 1;
            return result;
        }

        /// <summary>apply 合并：每个属性各自取"最后给出的非 null 值"。
        /// null 语义是"这次没碰这个属性"，所以要跳过而不是覆盖成 null。</summary>
        public static ApplySpec MergeApply(ApplySpec earlier, ApplySpec later)
        {
            if (earlier == null) return later;
            if (later == null) return earlier;

            return new ApplySpec
            {
                EastAsiaFont = later.EastAsiaFont ?? earlier.EastAsiaFont,
                LatinFont = later.LatinFont ?? earlier.LatinFont,
                SizePt = later.SizePt ?? earlier.SizePt,
                Bold = later.Bold ?? earlier.Bold,
                Color = later.Color ?? earlier.Color
            };
        }

        // ---- 比较辅助 ----

        private static bool EqIgnoreCase(string a, string b)
        {
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        private static bool EqNullable(double? a, double? b)
        {
            if (!a.HasValue || !b.HasValue) return a.HasValue == b.HasValue;
            return Math.Abs(a.Value - b.Value) < 0.01;
        }

        private static bool EqNullable(bool? a, bool? b)
        {
            return a.HasValue == b.HasValue && (!a.HasValue || a.Value == b.Value);
        }

        private static bool SameIgnore(List<string> a, List<string> b)
        {
            var setA = NormalizeIgnore(a);
            var setB = NormalizeIgnore(b);
            return setA.SetEquals(setB);
        }

        private static HashSet<string> NormalizeIgnore(List<string> list)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (list != null) foreach (string s in list) if (!string.IsNullOrEmpty(s)) set.Add(s);
            return set;
        }
    }
}
