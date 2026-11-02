using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using TypeRight.Formatting;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>公式化后同类匹配保真取证（P9）：
    /// ① 转换前后同族命中数对比（用户实测：公式化后 24 处同类只剩 7 处——签名被公式区污染）
    /// ② 字号批量应用零漏网验证（字号护栏回写旧值 bug：改完视觉无变化）
    /// ③ 公式转换耗时观测（草稿文档复用 vs 每公式开合窗口）。</summary>
    internal static class DiagP9
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.OpenCopy("p9");
            try
            {
                // ── ① 转换前：找一个"文字+公式候选"的正文档落样本 ──
                Paragraph body = FindBodySample(doc);
                if (body == null) { Console.WriteLine("!! 找不到正文样本"); return; }
                int bodyStart = body.Range.Start;
                int bodyOffset = 3;

                var pre = SimilarSelector.MatchUnits(doc, body, bodyOffset, FormatSignature.IgnoreDims.None);
                Console.WriteLine("[前] 正文样本 @段首=" + bodyStart + " sig=" + pre.SampleSignature);
                Console.WriteLine("[前] 命中 " + pre.Count + " 处 / " + pre.Duration.TotalMilliseconds.ToString("0") + "ms");

                // 全文公式区（转换前应为 null 或很少）
                List<int[]> zones0 = SimilarSelector.MathZoneRanges(doc);
                Console.WriteLine("[前] 公式区数=" + (zones0 == null ? 0 : zones0.Count));

                // ── ② 全文公式转换（用户路径），观测耗时 ──
                var watch = System.Diagnostics.Stopwatch.StartNew();
                var converter = new FormulaConverter(doc, null, FormulaConverter.ResolveMml2OmmlPath(env.App));
                FormulaStats stats = converter.ConvertDocument();
                watch.Stop();
                Console.WriteLine("[转换] " + stats.SummaryText() + "（墙钟 " + (watch.ElapsedMilliseconds / 1000.0).ToString("0.0") + "s，DCOM 环境偏慢）");
                if (stats.TimingText != null) Console.WriteLine("[转换] " + stats.TimingText);

                List<int[]> zones1 = SimilarSelector.MathZoneRanges(doc);
                Console.WriteLine("[后] 公式区数=" + (zones1 == null ? 0 : zones1.Count));

                // ── ③ 转换后：同样本再匹配（公式区感知后应恢复同族命中） ──
                // Word 的段落 COM 对象是"活"的：转换增删字符后其 Range 自动跟随，直接复用
                var post = SimilarSelector.MatchUnits(doc, body, bodyOffset, FormatSignature.IgnoreDims.None);
                Console.WriteLine("[后] 同样本 sig=" + post.SampleSignature);
                Console.WriteLine("[后] 命中 " + post.Count + " 处 / " + post.Duration.TotalMilliseconds.ToString("0") + "ms");
                Console.WriteLine("  [耗时分解] 枚举=" + (FormatSignatureReader.TicksEnumerator / 10000) + "ms 区间=" + (FormatSignatureReader.TicksRange / 10000)
                    + "ms 文本=" + (FormatSignatureReader.TicksText / 10000) + "ms 探针=" + (FormatSignatureReader.TicksProbe / 10000)
                    + "ms 字体=" + (FormatSignatureReader.TicksFont / 10000) + "ms 对齐=" + (FormatSignatureReader.TicksAlign / 10000) + "ms");

                // 掉队归因：前命中 vs 后命中逐段比对
                try
                {
                    var postStarts = new HashSet<int>();
                    foreach (MatchUnit u in post.Units) { try { postStarts.Add(u.Paragraph.Range.Start); } catch { } }
                    int pureMath = 0, mismatched = 0;
                    foreach (MatchUnit u in pre.Units)
                    {
                        int s;
                        try { s = u.Paragraph.Range.Start; } catch { continue; }
                        if (postStarts.Contains(s)) continue;
                        FormatSignature after = FormatSignatureReader.Read(u.Paragraph, zones1);
                        if (after.IsEmpty) { pureMath++; continue; }
                        mismatched++;
                        if (mismatched <= 4)
                        {
                            Console.WriteLine("  [掉队] 段首=" + s + " 后签名=" + after
                                + (after.Matches(post.SampleSignature, FormatSignature.IgnoreDims.None) ? "（匹配?!）" : ""));
                            // 深挖：边界检测 + 引导句单元签名
                            try
                            {
                                int b = SimilarSelector.DetectLeadBoundary(u.Paragraph, after, zones1);
                                Console.WriteLine("    边界=" + b + "（引导=" + (b < 0 ? -1 : b - s) + " 字符）"
                                    + (b >= 0 ? " 引导sig=" + FormatSignatureReader.ReadRange(doc.Range(s, b), zones1) : ""));
                            }
                            catch (Exception ex) { Console.WriteLine("    边界检测 EX: " + ex.Message); }
                        }
                    }
                    Console.WriteLine("  [归因] 纯公式段退出 " + pureMath + " 处，真掉队 " + mismatched + " 处");
                }
                catch { }

                // ── ④ 字号批量应用零漏网（护栏回写 bug 验证） ──
                double target = 12.5;
                var applied = BatchFormatter.ApplyUnits(env.App, doc, post.Units,
                    new FormatChanges { SizePt = target }, "P9 字号验证");
                int leaks = 0;
                List<int[]> zones2 = SimilarSelector.MathZoneRanges(doc);
                foreach (MatchUnit u in post.Units)
                {
                    try
                    {
                        FormatSignature s2 = FormatSignatureReader.Read(u.Paragraph, zones2);
                        if (double.IsNaN(s2.SizePt) || Math.Abs(s2.SizePt - target) >= 0.01)
                        {
                            leaks++;
                            if (leaks <= 5) Console.WriteLine("  [漏] 段首=" + u.Paragraph.Range.Start + " size=" + s2.SizePt);
                        }
                    }
                    catch { }
                }
                Console.WriteLine("应用 " + applied.Total + " 单元（改动 " + applied.Changed + "），首锚字号漏网 " + leaks + " 处（期望 0）");

                // 撤销字号改动，保持文档只读关闭
                try { doc.Undo(50); } catch { }
            }
            finally { try { doc.Close(false); } catch { } }
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
                    foreach (int[] t in tables)
                        if (start >= t[0] && start < t[1]) goto next;
                    string text = (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v');
                    if (text.Trim().Length < 30) continue;
                    bool hasSeg = TypeRight.Formulas.MathSegmenter.FindSegments(text).Count > 0;
                    if (hasSeg) return p;
                    next: ;
                }
                catch { }
            }
            return null;
        }

        private static Paragraph FindParagraphByStart(Document doc, int start)
        {
            foreach (Paragraph p in doc.Paragraphs)
            {
                try { if (p.Range.Start == start) return p; } catch { }
            }
            return null;
        }
    }
}
