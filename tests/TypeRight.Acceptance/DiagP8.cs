using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>高亮越界取证：正文单元匹配后，逐段检查黄色区间是否越过引导边界；
    /// 以及 Clear() 后是否彻底还原。</summary>
    internal static class DiagP8
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.App.Documents.Open(env.OriginalDocx, false, true);
            try
            {
                Paragraph box = doc.Paragraphs[3];
                int paraStart = box.Range.Start;
                var sig = FormatSignatureReader.Read(box);
                int boundary = SimilarSelector.DetectLeadBoundary(box, sig);
                Console.WriteLine("★框#3: start=" + paraStart + " 边界=" + boundary + "（引导=" + (boundary - paraStart) + "字符）");

                // 光标放正文（边界+5），单元匹配 + 单元高亮
                var result = SimilarSelector.MatchUnits(doc, box, (boundary - paraStart) + 5, FormatSignature.IgnoreDims.None);
                Console.WriteLine("正文样本命中 " + result.Count + " 处，样本单元=" + (result.SampleIsLead ? "引导句" : "正文"));
                var session = new SimilarSelector.HighlightSession(doc);
                int applied = session.ApplyUnits(result);
                Console.WriteLine("高亮写入 " + applied + " 处");

                // 检查 ★ 框#3 内黄色区间的分布（按 8 字符步进探测）
                DumpYellow(doc, box, paraStart, boundary);

                // 再检查所有命中单元的区间是否越界（Start 落在引导区内 = 越界）
                int leak = 0;
                foreach (var u in result.Units)
                {
                    try
                    {
                        int myBoundary = SimilarSelector.DetectLeadBoundary(u.Paragraph, FormatSignatureReader.Read(u.Paragraph));
                        if (myBoundary > 0 && u.Start < myBoundary && !u.IsLead) { leak++; Console.WriteLine("  [越界] 单元 Start=" + u.Start + " < 边界 " + myBoundary); }
                    }
                    catch { }
                }
                Console.WriteLine("单元区间越界数=" + leak);

                // 变体取证：高亮区间是否含结束标记导致整段扩散
                Console.WriteLine("== 变体A: [boundary, End-1) 排除段落/单元格标记 ==");
                TestVariant(doc, box, paraStart, boundary, box.Range.End - 1);
                Console.WriteLine("== 变体B: [boundary, End-2) 多退一格 ==");
                TestVariant(doc, box, paraStart, boundary, box.Range.End - 2);
                Console.WriteLine("== 变体C: [boundary+1, End-1) 起点也退一格 ==");
                TestVariant(doc, box, paraStart, boundary + 1, box.Range.End - 1);
            }
            finally { try { doc.Close(false); } catch { } }
        }

        private static void TestVariant(Document doc, Paragraph p, int paraStart, int boundary, int endPos)
        {
            try
            {
                // 先确保无黄
                try { p.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight; } catch { }
                var r = doc.Range(boundary, endPos);
                r.HighlightColorIndex = WdColorIndex.wdYellow;
                DumpYellow(doc, p, paraStart, boundary);
                try { p.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight; } catch { }
            }
            catch (Exception ex) { Console.WriteLine("  EX: " + ex.Message); }
        }

        private static void DumpYellow(Document doc, Paragraph p, int paraStart, int boundary, bool afterClear = false)
        {
            string tag = afterClear ? "Clear后" : "Clear前";
            var marks = new System.Text.StringBuilder();
            int len = p.Range.End - paraStart;
            int step = Math.Max(4, len / 24);
            for (int off = 0; off < len - 1; off += step)
            {
                try
                {
                    var idx = doc.Range(paraStart + off, Math.Min(paraStart + off + step, p.Range.End - 1)).HighlightColorIndex;
                    marks.Append(idx == WdColorIndex.wdYellow ? 'Y' : (idx == WdColorIndex.wdNoHighlight ? '.' : '?'));
                }
                catch { marks.Append('x'); }
            }
            int leadChars = Math.Max(1, (boundary - paraStart) / step);
            string leadZone = new string('L', leadChars);
            Console.WriteLine("  [" + tag + "] 黄色分布: " + marks + "  (前" + leadChars + "格=引导区)");
        }
    }
}
