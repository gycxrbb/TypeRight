using System;
using System.Text.RegularExpressions;
using NetOffice.WordApi;

namespace TypeRight.Acceptance
{
    /// <summary>InsertFile 最小复现：一段"A x^2 B"转换后看段落/文本/公式落点。</summary>
    internal static class DiagMini
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            string xsl = TypeRight.Formatting.FormulaConverter.ResolveMml2OmmlPath(env.App);

            Document doc = env.App.Documents.Add();
            try
            {
                var rng = doc.Range(0, 0);
                rng.InsertAfter("A x^2 B");
                Console.WriteLine("转换前 段落数=" + doc.Paragraphs.Count + " 文本=[" + doc.Content.Text.TrimEnd('\r') + "]");

                var conv = new TypeRight.Formatting.FormulaConverter(doc, null, xsl);
                var st = conv.ConvertDocument();
                Console.WriteLine("统计: " + st.SummaryText());

                Console.WriteLine("转换后 段落数=" + doc.Paragraphs.Count + " oMath=" + doc.OMaths.Count);
                foreach (Paragraph p in doc.Paragraphs)
                {
                    Console.WriteLine("  段 [" + (p.Range.Text ?? "").Replace("\r", "¶").Replace("\a", "§") + "]");
                }
                doc.SaveAs2(env.CopyPath("mini"));
                doc.Close(false);
            }
            finally { try { doc.Close(false); } catch { } }

            using (var xml = DocXml.Load(env.CopyPath("mini")))
            {
                Console.WriteLine("文件级: 段=" + xml.NonEmptyParagraphCount() + " oMath=" + xml.OMathCount() +
                    " sSup=" + xml.MathElementCount("sSup"));
                string t = DocXml.PlainText(xml.Root) + "|" + DocXml.MathText(xml.Root);
                Console.WriteLine("文件文本: [" + t + "]");
            }
        }
    }
}
