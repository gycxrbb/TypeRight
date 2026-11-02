using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using TypeRight.Formatting;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>TC-21（P9 修复验收）：全局公式化之后——
    /// ① 同类选择匹配不退化（签名公式区感知：锚点不再读到 OMML 区内部的数学字体值；
    ///    用户实测公式化后 24 处同类只剩 7 处）
    /// ② 批量改字号零漏网（字号护栏回写旧值 bug：改完"实际改动 24"但视觉无变化）。</summary>
    public static class CasesE
    {
        public static void TC21(TestEnv env)
        {
            env.Begin("TC-21 公式化后匹配保真 + 批量字号零漏网（P9）");
            Document doc = env.OpenCopy("tc21");
            try
            {
                // ── ① 转换前基线：找一个"文字+公式候选"的正文段 ──
                Paragraph body = FindBodySample(doc);
                if (!Found(env, body != null, "找到文字+公式正文样本段")) { doc.Close(false); return; }
                var pre = SimilarSelector.MatchUnits(doc, body, 3, FormatSignature.IgnoreDims.None);
                Console.WriteLine("  [前] sig=" + pre.SampleSignature + " 命中 " + pre.Count);
                if (!Found(env, pre.Count >= 2, "转换前同族命中 ≥ 2")) { doc.Close(false); return; }
                double preSize = pre.SampleSignature.SizePt;

                // ── ② 全文公式转换（Ribbon"公式转换"同路径） ──
                var stats = new FormulaConverter(doc, null, FormulaConverter.ResolveMml2OmmlPath(env.App)).ConvertDocument();
                Console.WriteLine("  [转换] " + stats.SummaryText());
                env.Check(stats.Converted >= 100, "公式转换数量级正常（≥100 处）", ">=100", stats.Converted.ToString(), null);

                // ── ③ 转换后同样本匹配：签名不被公式区污染（段落 COM 对象是活的，直接复用） ──
                var post = SimilarSelector.MatchUnits(doc, body, 3, FormatSignature.IgnoreDims.None);
                Console.WriteLine("  [后] sig=" + post.SampleSignature + " 命中 " + post.Count);

                // 中段采样一致性：公式区感知下，"文字+公式"正文段的中段锚点应与首锚一致
                // （修复前中段读到公式区内部的 10.5pt/B0 之类的数学字体值）
                env.Check(post.SampleSignature.MiddleUniform(), "样本中段采样无公式区污染", "一致",
                    post.SampleSignature.MiddleUniform() ? "一致" : "中段异于首锚", post.SampleSignature.ToString());

                // 同族命中不退化：允许"整段皆公式"的成员退出（转后无可见文字锚点）
                var zones = SimilarSelector.MathZoneRanges(doc);
                int expectedDrop = 0;
                foreach (MatchUnit u in pre.Units)
                {
                    try
                    {
                        if (FormatSignatureReader.Read(u.Paragraph, zones).IsEmpty) expectedDrop++;
                    }
                    catch { expectedDrop++; }
                }
                Console.WriteLine("  [允许退出] 整段皆公式的同族成员 " + expectedDrop + " 处");
                env.Check(post.Count >= pre.Count - expectedDrop, "公式化后同族命中不退化",
                    "≥" + (pre.Count - expectedDrop), post.Count.ToString(),
                    "前=" + pre.Count + " 后=" + post.Count + " 允许退出=" + expectedDrop);

                // ── ④ 批量字号零漏网 ──
                double target = Math.Abs(preSize - 12.5) < 0.01 ? 14 : 12.5;
                var applied = BatchFormatter.ApplyUnits(env.App, doc, post.Units,
                    new FormatChanges { SizePt = target }, "TC-21 字号验证");
                int leaks = 0;
                string firstLeak = null;
                foreach (MatchUnit u in post.Units)
                {
                    try
                    {
                        int end = u.End > u.Start + 1 ? u.End - 1 : u.End;
                        FormatSignature s2 = FormatSignatureReader.ReadRange(doc.Range(u.Start, end), zones);
                        if (double.IsNaN(s2.SizePt) || Math.Abs(s2.SizePt - target) >= 0.05)
                        {
                            leaks++;
                            if (firstLeak == null) firstLeak = "段首=" + u.Start + " size=" + s2.SizePt;
                        }
                    }
                    catch { leaks++; }
                }
                env.Check(leaks == 0, "批量字号应用零漏网（每单元首锚=目标字号）", "0",
                    leaks.ToString(), firstLeak);
                Console.WriteLine("  [字号] 应用 " + applied.Total + " 单元（改动 " + applied.Changed + "）→ 目标 " + target + "pt，漏网 " + leaks);
            }
            finally { try { doc.Close(false); } catch { } }
        }

        private static bool Found(TestEnv env, bool cond, string what)
        {
            env.Check(cond, what, "满足", cond ? "满足" : "不满足", null);
            return cond;
        }

        /// <summary>找"文字+公式候选"正文段：非表格、可见文本 ≥30 字符、含公式候选片段。</summary>
        private static Paragraph FindBodySample(Document doc)
        {
            List<int[]> tables = SimilarSelector.TableRangesOf(doc);
            foreach (Paragraph p in doc.Paragraphs)
            {
                try
                {
                    int start = p.Range.Start;
                    bool inTable = false;
                    foreach (int[] t in tables)
                        if (start >= t[0] && start < t[1]) { inTable = true; break; }
                    if (inTable) continue;
                    string text = (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v');
                    if (text.Trim().Length < 30) continue;
                    if (TypeRight.Formulas.MathSegmenter.FindSegments(text).Count > 0) return p;
                }
                catch { }
            }
            return null;
        }

    }
}