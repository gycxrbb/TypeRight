using System;
using System.Collections.Generic;
using System.Linq;
using TypeRight.Roles;
using TypeRight.Scanning;

namespace TypeRight.Classification
{
    /// <summary>分类上下文：需要全文统计才能判断的信息（正文基准字号等）。</summary>
    public sealed class ClassifierContext
    {
        /// <summary>全文正文聚类字号（长段落字号众数），NaN=样本不足。</summary>
        public double BodyFontSize = double.NaN;

        /// <summary>R11 聚类结果：签名 → 角色。null 表示尚未计算。</summary>
        public Dictionary<string, ParagraphRole> ClusterRoles;
    }

    /// <summary>单条规则：命中返回角色，未命中返回 null。</summary>
    public sealed class RoleRule
    {
        public int Priority;
        public string Id;
        public string Description;
        public Func<ParagraphFeatures, ClassifierContext, ParagraphRole?> Match;

        public override string ToString() { return "R" + Priority + " " + Description; }
    }

    /// <summary>
    /// 规则优先级表（PRD §4.2 第二段），严格按序，命中即停。
    /// R1–R10 为高置信；R11 聚类与 R12 兜底为低置信。
    /// </summary>
    public static class RuleTable
    {
        public const int MinHeadingLength = 0;
        public const int MaxHeadingLength = 40;

        public const int PatternNone = 0;
        public const int PatternChapter = 1;
        public const int PatternNumeric2 = 2;
        public const int PatternNumeric3 = 3;
        public const int PatternChinese = 4;
        public const int PatternCircled = 5;
        public const int PatternBracket = 6;

        /// <summary>高置信规则（R1–R10），逐段独立可判。</summary>
        public static readonly List<RoleRule> HighConfidenceRules = new List<RoleRule>
        {
            new RoleRule
            {
                Priority = 1, Id = "R1", Description = "空段 / 纯图片段 / 域代码段 → skip",
                Match = (f, ctx) => f.IsEmpty || f.HasInlineImage || f.HasFields ? (ParagraphRole?)ParagraphRole.Skip : null
            },
            new RoleRule
            {
                Priority = 2, Id = "R2", Description = "表格内：首行加粗→表头，其余→表格文字",
                Match = (f, ctx) =>
                {
                    if (!f.InTable) return null;
                    return f.IsFirstTableRow && (f.Bold || f.BoldMixed) ? ParagraphRole.TableHeader : ParagraphRole.TableText;
                }
            },
            new RoleRule
            {
                Priority = 3, Id = "R3", Description = "已有 Heading 样式 → 对应级别",
                Match = (f, ctx) => f.HeadingStyleLevel >= 1 && f.HeadingStyleLevel <= 3
                    ? (ParagraphRole)(ParagraphRole.Heading1 + (f.HeadingStyleLevel - 1))
                    : null
            },
            new RoleRule
            {
                Priority = 4, Id = "R4", Description = "全文第一个非空段 且 字号≥正文+4pt 且 加粗 → 主标题",
                Match = (f, ctx) =>
                {
                    if (!f.IsFirstNonEmpty || f.InTable || f.BoldMixed) return null;
                    if (double.IsNaN(ctx.BodyFontSize) || double.IsNaN(f.FontSize)) return null;
                    return f.FontSize >= ctx.BodyFontSize + 4 && f.Bold ? (ParagraphRole?)ParagraphRole.Title : null;
                }
            },
            new RoleRule
            {
                Priority = 5, Id = "R5", Description = "第X章/第X部分/模块X 且 长度≤40 → 一级标题",
                Match = (f, ctx) => f.NumberingPattern == PatternChapter && f.Text.Length <= MaxHeadingLength
                    ? (ParagraphRole?)ParagraphRole.Heading1 : null
            },
            new RoleRule
            {
                Priority = 6, Id = "R6", Description = "X.X.X 且 长度≤40 → 三级标题",
                Match = (f, ctx) => f.NumberingPattern == PatternNumeric3 && f.Text.Length <= MaxHeadingLength
                    ? (ParagraphRole?)ParagraphRole.Heading3 : null
            },
            new RoleRule
            {
                Priority = 7, Id = "R7", Description = "X.X 且 长度≤40 → 二级标题",
                Match = (f, ctx) => f.NumberingPattern == PatternNumeric2 && f.Text.Length <= MaxHeadingLength
                    ? (ParagraphRole?)ParagraphRole.Heading2 : null
            },
            new RoleRule
            {
                Priority = 8, Id = "R8", Description = "图注/表注模式 → 图注",
                Match = (f, ctx) => f.CaptionPattern ? (ParagraphRole?)ParagraphRole.Caption : null
            },
            new RoleRule
            {
                Priority = 9, Id = "R9", Description = "Word 自动编号/项目符号 → 列表项",
                Match = (f, ctx) => f.HasAutoNumbering ? (ParagraphRole?)ParagraphRole.ListItem : null
            },
            new RoleRule
            {
                Priority = 10, Id = "R10", Description = "独立成行 且 数学密度>0.7 且 含= → 公式行",
                Match = (f, ctx) => f.StandaloneMath && f.MathDensity > 0.7 && f.HasEquals
                    ? (ParagraphRole?)ParagraphRole.FormulaLine : null
            }
        };

        /// <summary>对单个段落跑高置信规则；未命中返回 null（交给 R11/R12）。</summary>
        public static ParagraphRole? EvaluateHighConfidence(ParagraphFeatures f, ClassifierContext ctx)
        {
            foreach (var rule in HighConfidenceRules)
            {
                var role = rule.Match(f, ctx);
                if (role.HasValue)
                {
                    f.RulePriority = rule.Priority;
                    f.HighConfidence = true;
                    return role.Value;
                }
            }
            return null;
        }
    }

    /// <summary>
    /// R11 格式聚类兜底：按（字号, 加粗, 颜色, 对齐）签名聚类；
    /// 占比最大的簇 = body；比 body 字号大 ≥1.5pt 且加粗的簇按字号降序补位 heading1/2/3（每级限一簇）。
    /// </summary>
    public static class ClusterAnalyzer
    {
        public static string SignatureOf(ParagraphFeatures f)
        {
            string size = double.IsNaN(f.FontSize) ? "?" : Math.Round(f.FontSize, 1).ToString("0.#");
            string bold = (f.Bold || f.BoldMixed) ? "B" : "b";
            string color = f.ColorRgb >= 0 ? f.ColorRgb.ToString("X6") : "?";
            return size + "|" + bold + "|" + color + "|" + f.Alignment;
        }

        /// <summary>返回 签名 → 角色 映射（R11 判定结果）。</summary>
        public static Dictionary<string, ParagraphRole> Analyze(IList<ParagraphFeatures> undecided)
        {
            var result = new Dictionary<string, ParagraphRole>();
            if (undecided == null || undecided.Count == 0) return result;

            // 簇统计：签名 → 段数、代表字号
            var clusters = new Dictionary<string, List<ParagraphFeatures>>();
            foreach (var f in undecided)
            {
                string sig = SignatureOf(f);
                f.ClusterSignature = sig;
                List<ParagraphFeatures> list;
                if (!clusters.TryGetValue(sig, out list)) { list = new List<ParagraphFeatures>(); clusters[sig] = list; }
                list.Add(f);
            }

            // body 簇 = 段数占比最大
            string bodySig = null;
            int bodyCount = -1;
            foreach (var kv in clusters)
            {
                if (kv.Value.Count > bodyCount)
                {
                    bodyCount = kv.Value.Count;
                    bodySig = kv.Key;
                }
            }
            foreach (var sig in clusters.Keys) result[sig] = ParagraphRole.Body;
            if (bodySig == null) return result;

            double bodySize = SizeOf(clusters[bodySig]);

            // heading 候选簇：字号 ≥ body+1.5 且加粗；按字号降序取前 3
            var headingCandidates = clusters
                .Where(kv => kv.Key != bodySig)
                .Select(kv => new { Sig = kv.Key, Size = SizeOf(kv.Value), Bold = kv.Key.Contains("|B|") })
                .Where(x => !double.IsNaN(x.Size) && x.Bold && x.Size >= bodySize + 1.5)
                .OrderByDescending(x => x.Size)
                .Take(3)
                .ToList();

            var headingRoles = new[] { ParagraphRole.Heading1, ParagraphRole.Heading2, ParagraphRole.Heading3 };
            for (int i = 0; i < headingCandidates.Count && i < headingRoles.Length; i++)
                result[headingCandidates[i].Sig] = headingRoles[i];

            return result;
        }

        private static double SizeOf(List<ParagraphFeatures> cluster)
        {
            // 簇内签名相同 → 字号一致；NaN 的簇返回 NaN
            return cluster[0].FontSize;
        }
    }

    /// <summary>
    /// 两段式编排：① 提取特征（DocumentScanner）→ 计算正文基准字号 → 逐段跑 R1–R10；
    /// ② 对未命中的段落做 R11 聚类，R12 兜底 → body。
    /// </summary>
    public static class RoleClassifier
    {
        /// <summary>计算正文基准字号：长度≥20 的非表格段的字号众数（无样本返回 NaN）。</summary>
        public static double ComputeBodyFontSize(IList<ParagraphFeatures> features)
        {
            var counts = new Dictionary<double, int>();
            foreach (var f in features)
            {
                if (f.InTable || f.IsEmpty || double.IsNaN(f.FontSize)) continue;
                if (f.Text.Length < 20) continue;
                int c;
                counts[f.FontSize] = counts.TryGetValue(f.FontSize, out c) ? c + 1 : 1;
            }
            double best = double.NaN;
            int bestCount = 0;
            foreach (var kv in counts)
            {
                if (kv.Value > bestCount)
                {
                    bestCount = kv.Value;
                    best = kv.Key;
                }
            }
            return best;
        }

        public static void Classify(IList<ParagraphFeatures> features)
        {
            // 全文第一个非空段（表格外的第一个有可见文本的段落）
            bool firstSeen = false;
            foreach (var f in features)
            {
                if (!firstSeen && !f.IsEmpty && !f.InTable && !f.IsTocField)
                {
                    f.IsFirstNonEmpty = true;
                    firstSeen = true;
                }
            }

            var ctx = new ClassifierContext { BodyFontSize = ComputeBodyFontSize(features) };

            var undecided = new List<ParagraphFeatures>();
            foreach (var f in features)
            {
                var role = RuleTable.EvaluateHighConfidence(f, ctx);
                if (role.HasValue)
                {
                    f.Role = role.Value;
                }
                else
                {
                    f.Role = ParagraphRole.Body; // R12 兜底，可能被 R11 覆盖
                    f.RulePriority = 12;
                    f.HighConfidence = false;
                    undecided.Add(f);
                }
            }

            var clusterRoles = ClusterAnalyzer.Analyze(undecided);
            foreach (var f in undecided)
            {
                if (f.ClusterSignature != null)
                {
                    ParagraphRole role;
                    if (clusterRoles.TryGetValue(f.ClusterSignature, out role) && role != ParagraphRole.Body)
                    {
                        f.Role = role;
                        f.RulePriority = 11;
                    }
                }
            }
        }
    }
}
