using System;
using System.Text;
using NetOffice.WordApi;

namespace TypeRight.Acceptance
{
    /// <summary>InsertXML 注入 OMML 探针。</summary>
    internal static class DiagInsert
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = Encoding.UTF8;

            // ── A. BuildUp 变体（空格 / U+2215 / 完整参数式）──
            string[] variants =
            {
                "A1 a / b",
                "A2 a∕b",
                "A3 (a)/(b)",
                "A4 x = (-b±√(b^2-4ac))/(2a)",
                "A5 √(b^2-4ac)",
            };
            Document probe = env.App.Documents.Add();
            try
            {
                var rng = probe.Range(0, 0);
                foreach (string v in variants) { rng.InsertAfter(v); rng.InsertParagraphAfter(); }
                var conv = new TypeRight.Formatting.FormulaConverter(probe, null, null);
                conv.ConvertDocument();
                probe.SaveAs2(env.CopyPath("ins-a"));
                probe.Close(false);
            }
            catch (Exception ex) { Console.WriteLine("A 外层异常 " + ex.Message); try { probe.Close(false); } catch { } }
            using (var xml = DocXml.Load(env.CopyPath("ins-a")))
            {
                Console.WriteLine("A(BuildUp变体): oMath=" + xml.OMathCount() + " m:f=" + xml.MathElementCount("f") +
                    " m:rad=" + xml.MathElementCount("rad") + " sSup=" + xml.MathElementCount("sSup"));
            }

            // ── B. OMathFunctions.Add 结构化构造 ──
            Document probe2 = env.App.Documents.Add();
            try
            {
                // B1: 区间含 "c/d" 文本 → Functions.Add(range, Fraction)
                var r1 = probe2.Range(0, 0);
                r1.InsertAfter("B1: c/d");
                probe2.SaveAs2(env.CopyPath("ins-b1")); // 占位
                var find1 = probe2.Content;
                find1.Find.Text = "c/d";
                find1.Find.Wrap = NetOffice.WordApi.Enums.WdFindWrap.wdFindStop;
                if (find1.Find.Execute())
                {
                    try
                    {
                        Range zone = probe2.OMaths.Add(find1);
                        var om = zone.OMaths[1];
                        var fn = om.Functions.Add(om.Range, NetOffice.WordApi.Enums.WdOMathFunctionType.wdOMathFunctionFrac);
                        Console.WriteLine("B1 Functions.Add(Fraction) OK: " + (fn != null));
                    }
                    catch (Exception ex) { Console.WriteLine("B1 异常: " + ex.Message); }
                }

                // B2: 空区建根式再填文本
                var r2 = probe2.Range(probe2.Content.End - 1, probe2.Content.End - 1);
                r2.InsertAfter("B2: kkk");
                var find2 = probe2.Content;
                find2.Find.Text = "kkk";
                find2.Find.Wrap = NetOffice.WordApi.Enums.WdFindWrap.wdFindStop;
                if (find2.Find.Execute())
                {
                    try
                    {
                        Range zone2 = probe2.OMaths.Add(find2);
                        var om2 = zone2.OMaths[1];
                        var fn2 = om2.Functions.Add(om2.Range, NetOffice.WordApi.Enums.WdOMathFunctionType.wdOMathFunctionRad);
                        Console.WriteLine("B2 Functions.Add(Rad) OK: " + (fn2 != null));
                    }
                    catch (Exception ex) { Console.WriteLine("B2 异常: " + ex.Message); }
                }

                probe2.SaveAs2(env.CopyPath("ins-b"));
                probe2.Close(false);
            }
            catch (Exception ex) { Console.WriteLine("B 外层异常 " + ex.Message); try { probe2.Close(false); } catch { } }
            using (var xml = DocXml.Load(env.CopyPath("ins-b")))
            {
                Console.WriteLine("B(Functions.Add): oMath=" + xml.OMathCount() + " m:f=" + xml.MathElementCount("f") +
                    " m:rad=" + xml.MathElementCount("rad"));
            }
        }
    }
}
