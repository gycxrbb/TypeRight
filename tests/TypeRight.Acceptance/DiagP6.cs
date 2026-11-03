using System;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>取证 ★ 导读框结构：是几段？段内 run 分布（哪些字符加粗蓝/哪些正文）？
    /// 以及样本签名命中的 9 段各是什么。</summary>
    internal static class DiagP6
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.App.Documents.Open(env.OriginalDocx, false, true);
            try
            {
                int idx = 0;
                foreach (Paragraph p in doc.Paragraphs)
                {
                    idx++;
                    string t = (p.Range.Text ?? "");
                    if (!t.Contains("降维打击")) continue;

                    Console.WriteLine("== 含'降维打击'的段 #" + idx + " 长度=" + t.Length + " ==");
                    Console.WriteLine("  段内含段落标记个数(应为0=单段): " + (t.Split('\r').Length - 1));
                    // run 级分布：前 6 个 run 的文本+格式
                    int runSeen = 0;
                    foreach (Range r in p.Range.Words)
                    {
                        if (runSeen >= 8) break;
                        string wt = (r.Text ?? "").TrimEnd('\r', '\a');
                        if (wt.Length == 0) continue;
                        var f = r.Font;
                        Console.WriteLine(string.Format("  word'{0}' {1}pt B={2} #{3:X6} EA={4}",
                            wt.Length > 6 ? wt.Substring(0, 6) : wt,
                            f.Size, f.Bold, (int)f.Color < 0 ? 0 : (int)f.Color, f.NameFarEast));
                        runSeen++;
                    }
                    var s = FormatSignatureReader.Read(p);
                    Console.WriteLine("  签名: " + s);
                }

                // 与样本同签名的 9 段都是什么
                Console.WriteLine("== 同签名(微软雅黑10.5B1首锚+宋体10尾锚)命中段 ==");
                idx = 0;
                foreach (Paragraph p in doc.Paragraphs)
                {
                    idx++;
                    var s = FormatSignatureReader.Read(p);
                    if (s.IsEmpty || s.EastAsiaFont == null || !s.EastAsiaFont.Contains("微软雅黑")) continue;
                    if (double.IsNaN(s.SizePt) || Math.Abs(s.SizePt - 10.5) > 0.01 || s.BoldRaw != -1 || s.ColorRgb != 0x1A365D) continue;
                    string t = (p.Range.Text ?? "").TrimEnd('\r', '\a');
                    if (t.Length > 30) t = t.Substring(0, 30) + "…";
                    Console.WriteLine("  #" + idx + " [" + t + "]");
                }
            }
            finally { try { doc.Close(false); } catch { } }
        }
    }
}
