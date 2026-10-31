using System;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>找出"首锚=微软雅黑10.5B1#1A365D、尾锚=宋体10pt#2D3748"的全部段落并打印文本，
    /// 取证用户截图里被判同类的"两个类别"到底是什么。</summary>
    internal static class DiagP5
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.App.Documents.Open(env.OriginalDocx, false, true);
            try
            {
                int idx = 0, hits = 0;
                foreach (Paragraph p in doc.Paragraphs)
                {
                    idx++;
                    var s = FormatSignatureReader.Read(p);
                    if (s.IsEmpty) continue;
                    bool firstIsHeadingStyle = s.EastAsiaFont != null && s.EastAsiaFont.Contains("微软雅黑")
                        && !double.IsNaN(s.SizePt) && Math.Abs(s.SizePt - 10.5) < 0.01
                        && s.BoldRaw == -1 && s.ColorRgb == 0x1A365D;
                    if (firstIsHeadingStyle)
                    {
                        hits++;
                        string t = (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v');
                        if (t.Length > 40) t = t.Substring(0, 40) + "…";
                        bool inTable = false;
                        try { inTable = SimilarSelector.IsInTable(doc, p); } catch { }
                        Console.WriteLine(string.Format("  #{0} 表内={1} 尾={2}pt B={3} #{4:X6} [{5}]",
                            idx, inTable,
                            double.IsNaN(s.LastSizePt) ? "?" : s.LastSizePt.ToString("0.#"),
                            s.LastBoldRaw == -1 ? 1 : (s.LastBoldRaw == 0 ? 0 : -1),
                            s.LastColorRgb < 0 ? 0 : s.LastColorRgb,
                            t.Replace("\r", "¶")));
                    }
                }
                Console.WriteLine("共 " + hits + " 段（首锚=微软雅黑10.5pt加粗#1A365D）");
            }
            finally { try { doc.Close(false); } catch { } }
        }
    }
}
