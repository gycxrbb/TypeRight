using System;
using NetOffice.WordApi;

namespace TypeRight.Acceptance
{
    /// <summary>
    /// 字符级粒度可行性探针（2026-11）。
    /// 回答："Word/WPS 能不能做到汉字归汉字、数字字母归数字字母"。
    /// 测三件事：
    ///  1. 单个字符的字体能否独立读取（Range(start,end) 只含该字）
    ///  2. 单个字符的字体能否独立写入（不影响相邻字符）
    ///  3. 逐字符扫描一份 178 字的真实段落要多少次 COM 往返 / 耗时多少
    /// </summary>
    internal static class DiagGranularity
    {
        public static int Run(TestEnv env, string path)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.App.Documents.Open(path, false, true, false);
            try
            {
                // 找第 5 段（正文，汉字114/数字49 混排）
                Paragraph target = null;
                foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
                {
                    string t = (p.Range.Text ?? "").Replace("\r", "").Replace("\a", "");
                    if (t.Length > 150) { target = p; break; }
                }
                if (target == null) { Console.WriteLine("[FAIL] 找不到长正文段"); return 1; }

                Range pr = target.Range;
                string text = (pr.Text ?? "").Replace("\r", "").Replace("\a", "");
                Console.WriteLine();
                Console.WriteLine("── 字符级粒度可行性探针 ──────────────────────");
                Console.WriteLine("目标段：长度 " + text.Length + " 字符");

                // ── 1. 单字符字体能否独立读取 ──
                Console.WriteLine();
                Console.WriteLine("[1] 单字符独立读取（前 12 字符）");
                int start = pr.Start;
                int shown = 0;
                for (int i = 0; i < text.Length && shown < 12; i++)
                {
                    char c = text[i];
                    if (c == '\r' || c == '\a') continue;
                    Range one = doc.Range(start + i, start + i + 1);
                    Font f = one.Font;
                    Console.WriteLine(string.Format("   [{0}] '{1}' Ea={2,-6} Ascii={3,-16} Size={4}",
                        i, c, f.NameFarEast, f.NameAscii, f.Size));
                    shown++;
                }

                // ── 2. 单字符字体能否独立写入 ──
                Console.WriteLine();
                Console.WriteLine("[2] 单字符独立写入测试（只在内存里改，不保存）");
                // 找第一个数字字符
                int digitIdx = -1;
                for (int i = 0; i < text.Length; i++)
                    if (text[i] >= '0' && text[i] <= '9') { digitIdx = i; break; }
                if (digitIdx < 0) { Console.WriteLine("   [SKIP] 段内无数字"); }
                else
                {
                    Range d = doc.Range(start + digitIdx, start + digitIdx + 1);
                    string beforeNeighbor = "";
                    try
                    {
                        Range prev = doc.Range(start + digitIdx - 1, start + digitIdx);
                        beforeNeighbor = prev.Font.NameFarEast ?? "";
                    }
                    catch { }

                    string orig = d.Font.NameFarEast ?? "";
                    d.Font.NameFarEast = "黑体";                  // 故意改成明显不同的字体
                    string after = d.Font.NameFarEast ?? "";
                    string neighborAfter = "";
                    try
                    {
                        Range prev = doc.Range(start + digitIdx - 1, start + digitIdx);
                        neighborAfter = prev.Font.NameFarEast ?? "";
                    }
                    catch { }

                    bool isolated = (after == "黑体") && (neighborAfter == beforeNeighbor);
                    Console.WriteLine(string.Format("   数字 '{0}' 原位={1} 改为黑体后={2}", text[digitIdx], orig, after));
                    Console.WriteLine(string.Format("   相邻字符字体：改前={0} 改后={1}", beforeNeighbor, neighborAfter));
                    Console.WriteLine("   → 单字符独立写入 " + (isolated ? "可行 ✓" : "不可行 ✗"));

                    d.Font.NameFarEast = orig;                     // 还原
                }

                // ── 3. 逐字符扫描成本 ──
                Console.WriteLine();
                Console.WriteLine("[3] 逐字符扫描成本（同一段全量）");
                var sw = System.Diagnostics.Stopwatch.StartNew();
                int visits = 0;
                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];
                    if (c == '\r' || c == '\a') continue;
                    Range one = doc.Range(start + i, start + i + 1);
                    Font f = one.Font;
                    // 模拟读 3 个属性
                    string ea = f.NameFarEast;
                    string asc = f.NameAscii;
                    float sz = f.Size;
                    visits++;
                    if (ea == null || asc == null || sz < 0) Console.Write("");
                }
                sw.Stop();
                Console.WriteLine(string.Format("   {0} 字符 × (造Range+读3属性) = {1} ms，平均 {2:0.00} ms/字符",
                    visits, sw.ElapsedMilliseconds, (double)sw.ElapsedMilliseconds / Math.Max(1, visits)));
                Console.WriteLine(string.Format("   推算 1000 字文档：{0:0} 秒", sw.ElapsedMilliseconds / 1000.0 / Math.Max(1, visits) * 1000));
            }
            finally { try { doc.Close(false); } catch { } }
            return 0;
        }
    }
}
