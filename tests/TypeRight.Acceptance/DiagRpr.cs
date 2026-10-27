using System;
using System.Text;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Acceptance
{
    /// <summary>TC-20 定位探针：公式转换前后表格 run 的 rPr 逐项对比。</summary>
    internal static class DiagRpr
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.OpenCopy("rpr-probe");
            try
            {
                Paragraph star = env.FindParagraph(doc, "★");
                if (star == null) { Console.WriteLine("未找到★段"); return; }

                Console.WriteLine("── 转换前 ★ 段 run 快照 ──");
                var before = DumpRuns(doc, star);

                var conv = new TypeRight.Formatting.FormulaConverter(doc, null, null);
                var st = conv.ConvertDocument();
                Console.WriteLine("转换: " + st.SummaryText());

                Console.WriteLine("── 转换后非公式 run 对比 ──");
                var after = DumpRuns(doc, star);

                int n = Math.Min(before.Count, after.Count);
                int diff = 0;
                for (int i = 0; i < n; i++)
                {
                    var b = before[i];
                    var a = after[i];
                    if (b.Item2 != a.Item2 || b.Item3 != a.Item3 || b.Item4 != a.Item4 || b.Item5 != a.Item5)
                    {
                        diff++;
                        if (diff <= 8)
                            Console.WriteLine(string.Format("  [{0}] '{1}'  前(sz={2},b={3},c={4},fe={5}) 后(sz={6},b={7},c={8},fe={9})",
                                i, b.Item1.Length > 8 ? b.Item1.Substring(0, 8) : b.Item1,
                                b.Item2, b.Item3, b.Item4, b.Item5, a.Item2, a.Item3, a.Item4, a.Item5));
                    }
                }
                Console.WriteLine("run 对比数=" + n + " 属性差异=" + diff);
            }
            finally { doc.Close(false); }
        }

        private static System.Collections.Generic.List<Tuple<string, float, int, int, string>> DumpRuns(Document doc, Paragraph p)
        {
            var list = new System.Collections.Generic.List<Tuple<string, float, int, int, string>>();
            var cells = new System.Collections.Generic.List<NetOffice.WordApi.Range>();
            try
            {
                foreach (var ch in p.Range.Characters)
                {
                    var r = ch as NetOffice.WordApi.Range;
                    if (r == null) continue;
                    string t = r.Text ?? "";
                    if (t.Length == 0) continue;
                    char c0 = t[0];
                    if (c0 == '\r' || c0 == '\a') continue;
                    int inMath = 0;
                    try { inMath = r.OMaths.Count; } catch { }
                    if (inMath > 0) continue;   // 公式区 run 不参与 rPr 保持断言
                    var f = r.Font;
                    list.Add(Tuple.Create(t, f.Size, f.Bold, (int)f.Color, f.NameFarEast ?? "?"));
                }
            }
            catch (Exception ex) { Console.WriteLine("DumpRuns 异常: " + ex.Message); }
            return list;
        }
    }
}
