using System;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>复现"标题与正文被判同类"：同字体同色、仅差加粗+字号的段落是否跨判。</summary>
    internal static class DiagP4
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.App.Documents.Add();
            try
            {
                // 3 段"标题"：16pt 加粗 #1A365D；6 段"正文"：10.5pt 不加粗 #1A365D；字体统一
                for (int i = 1; i <= 3; i++)
                {
                    Range h = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                    h.InsertAfter("【小标题" + i + "】测试标题段落\r");
                    h.Font.NameFarEast = "MS Mincho"; h.Font.NameAscii = "MS Mincho";
                    h.Font.Size = 16; h.Font.Bold = 1; h.Font.Color = (NetOffice.WordApi.Enums.WdColor)0x1A365D;
                }
                for (int i = 1; i <= 6; i++)
                {
                    Range b = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                    b.InsertAfter("这是正文段落" + i + "，普通字号普通粗细的内容文字。\r");
                    b.Font.NameFarEast = "MS Mincho"; b.Font.NameAscii = "MS Mincho";
                    b.Font.Size = 10.5f; b.Font.Bold = 0; b.Font.Color = (NetOffice.WordApi.Enums.WdColor)0x1A365D;
                }
                // 3 段"混合段"：行首符号 10.5pt 不加粗 + 其余文字 16pt 加粗（同色同字体）
                for (int i = 1; i <= 3; i++)
                {
                    Range m = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                    m.InsertAfter("☆混合段" + i + "：行首符号后是大标题文字\r");
                    m.Font.NameFarEast = "MS Mincho"; m.Font.NameAscii = "MS Mincho";
                    m.Font.Size = 10.5f; m.Font.Bold = 0; m.Font.Color = (NetOffice.WordApi.Enums.WdColor)0x1A365D;
                    Paragraph mp = doc.Paragraphs[doc.Paragraphs.Count - 1];
                    Range rest = doc.Range(mp.Range.Start + 1, mp.Range.End - 1);
                    rest.Font.Size = 16; rest.Font.Bold = 1;
                }
                // 两族"同首尾锚、不同中段"（用户实测投诉场景）：首=☆10.5B0 尾=12ptB0，中段一粗一细
                for (int i = 1; i <= 2; i++)
                {
                    Range m = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                    m.InsertAfter("★X族" + i + "：行首符号，中间粗体，结尾小字。\r");
                    m.Font.NameFarEast = "MS Mincho"; m.Font.NameAscii = "MS Mincho";
                    m.Font.Size = 10.5f; m.Font.Bold = 0; m.Font.Color = (NetOffice.WordApi.Enums.WdColor)0x1A365D;
                    Paragraph mp = doc.Paragraphs[doc.Paragraphs.Count - 1];
                    Range mid = doc.Range(mp.Range.Start + 4, mp.Range.End - 5);
                    mid.Font.Size = 16; mid.Font.Bold = 1;
                    Range tail = doc.Range(mp.Range.End - 5, mp.Range.End - 1);
                    tail.Font.Size = 12; tail.Font.Bold = 0;
                }
                for (int i = 1; i <= 2; i++)
                {
                    Range m = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                    m.InsertAfter("★Y族" + i + "：行首符号，中间正常，结尾小字。\r");
                    m.Font.NameFarEast = "MS Mincho"; m.Font.NameAscii = "MS Mincho";
                    m.Font.Size = 10.5f; m.Font.Bold = 0; m.Font.Color = (NetOffice.WordApi.Enums.WdColor)0x1A365D;
                    Paragraph mp = doc.Paragraphs[doc.Paragraphs.Count - 1];
                    Range tail = doc.Range(mp.Range.End - 5, mp.Range.End - 1);
                    tail.Font.Size = 12; tail.Font.Bold = 0;
                }

                Console.WriteLine("== 全部段落签名 ==");
                int idx = 0;
                foreach (Paragraph p in doc.Paragraphs)
                {
                    var s = FormatSignatureReader.Read(p);
                    Console.WriteLine("  #" + (idx++) + " [" + Trunc(p) + "] " + s);
                }

                // 标题段锚点级取证（抖动定位）：首/中×3/尾 的原始值
                for (int pi = 1; pi <= 3; pi++)
                {
                    var s = FormatSignatureReader.Read(doc.Paragraphs[pi]);
                    string mids = s.ExtraAnchors == null ? "null" : "";
                    if (s.ExtraAnchors != null)
                        foreach (var e in s.ExtraAnchors)
                            mids += string.Format("[{0}/{1} {2}pt B={3} #{4:X6}] ", e.Ea, e.Latin,
                                double.IsNaN(e.Size) ? -1 : e.Size, e.Bold, e.Color < 0 ? 0 : e.Color);
                    Console.WriteLine(string.Format("  [标题{0}] 首 B={1} size={2} #{3:X6} | 中 {4}| 尾 B={5} size={6}",
                        pi, s.BoldRaw, double.IsNaN(s.SizePt) ? -1 : s.SizePt, s.ColorRgb < 0 ? 0 : s.ColorRgb,
                        mids, s.LastBoldRaw, double.IsNaN(s.LastSizePt) ? -1 : s.LastSizePt));
                }

                // 混合段取证：首字符 vs 末字符的原始字号/加粗
                for (int pi = 9; pi <= 11; pi++)
                {
                    try
                    {
                        Paragraph p = doc.Paragraphs[pi];
                        int lastPos = p.Range.End - 1;
                        Range first = doc.Range(p.Range.Start, p.Range.Start + 1);
                        Range last = doc.Range(lastPos - 1, lastPos);
                        Console.WriteLine(string.Format("  [混合{0}] 首='{1}' {2}pt B={3} | 末='{4}' {5}pt B={6}",
                            pi, first.Text, first.Font.Size, first.Font.Bold, last.Text, last.Font.Size, last.Font.Bold));
                    }
                    catch (Exception ex) { Console.WriteLine("  [混合" + pi + "] EX " + ex.Message); }
                }

                Test(doc, env, 1, "小标题为样本");
                Test(doc, env, 5, "正文为样本");
                Test(doc, env, 10, "混合段为样本（首10.5/余16pt）");
                Test(doc, env, 13, "X族为样本（同首尾锚、中段16pt粗）");
                Test(doc, env, 15, "Y族为样本（同首尾锚、中段10.5细）");
            }
            finally { try { doc.Close(false); } catch { } }
        }

        private static void Test(Document doc, TestEnv env, int sampleIdx, string label)
        {
            Paragraph sample = doc.Paragraphs[sampleIdx];
            var sig = FormatSignatureReader.Read(sample);
            bool inTable = SimilarSelector.IsInTable(doc, sample);
            var result = SimilarSelector.Match(doc, sig, FormatSignature.IgnoreDims.None, inTable);
            Console.WriteLine("== " + label + " sig=" + sig + " → 命中 " + result.Count + " 段 ==");
            foreach (Paragraph m in result.Paragraphs) Console.WriteLine("   命中: [" + Trunc(m) + "]");
        }

        private static string Trunc(Paragraph p)
        {
            try
            {
                string t = (p.Range.Text ?? "").TrimEnd('\r', '\a');
                return t.Length > 12 ? t.Substring(0, 12) : t;
            }
            catch { return "?"; }
        }
    }
}
