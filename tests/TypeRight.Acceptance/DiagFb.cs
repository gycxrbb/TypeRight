using System;
using System.Linq;
using NetOffice.WordApi;

namespace TypeRight.Acceptance
{
    /// <summary>回退定位探针：全文转换 → 打印 MML 回退清单 + 文件级结构统计。</summary>
    internal static class DiagFb
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.OpenCopy("fb");
            try
            {
                var st = new TypeRight.Formatting.FormulaConverter(doc, null, TypeRight.Formatting.FormulaConverter.ResolveMml2OmmlPath(env.App)).ConvertDocument();
                Console.WriteLine("转换: " + st.SummaryText());
                var log = TypeRight.Formatting.FormulaConverter.FallbackLog;
                Console.WriteLine("回退总数: " + log.Count);
                foreach (string f in log.Take(25)) Console.WriteLine("  回退: " + f);

                doc.SaveAs2(env.CopyPath("fb-out"));
                doc.Close(false);
                using (var xml = DocXml.Load(env.CopyPath("fb-out")))
                {
                    Console.WriteLine("文件级: oMath=" + xml.OMathCount() +
                        " m:f=" + xml.MathElementCount("f") +
                        " m:rad=" + xml.MathElementCount("rad") +
                        " sSub=" + xml.MathElementCount("sSub") +
                        " sSup=" + xml.MathElementCount("sSup") +
                        " sSubSup=" + xml.MathElementCount("sSubSup") +
                        " 上下标合计=" + (xml.MathElementCount("sSub") + xml.MathElementCount("sSup") + xml.MathElementCount("sSubSup")) +
                        " Unicode残留=" + xml.UnicodeMathCharsInPlainText());
                }
            }
            finally { try { doc.Close(false); } catch { } }
        }
    }
}
