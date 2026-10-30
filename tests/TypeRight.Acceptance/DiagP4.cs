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

                Console.WriteLine("== 全部段落签名 ==");
                int idx = 0;
                foreach (Paragraph p in doc.Paragraphs)
                {
                    var s = FormatSignatureReader.Read(p);
                    Console.WriteLine("  #" + (idx++) + " [" + Trunc(p) + "] " + s);
                }

                Test(doc, env, 1, "小标题为样本");
                Test(doc, env, 5, "正文为样本");
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
