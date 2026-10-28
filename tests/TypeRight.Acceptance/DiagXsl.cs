using System;

namespace TypeRight.Acceptance
{
    /// <summary>XSL 转换输出探针：MathML 样例 → OMML 元素统计。</summary>
    internal static class DiagXsl
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            string xsl = TypeRight.Formatting.FormulaConverter.ResolveMml2OmmlPath(env.App);
            Console.WriteLine("XSL: " + xsl);
            if (string.IsNullOrEmpty(xsl)) return;

            string[] samples =
            {
                "x_1^2",
                "x_1",
                "a/b",
                "(-b±√(b^2-4ac))/(2a)",
                "x_1 + x_2",
                "(x_1 + x_2)^2",
                "(x_1 + x_2)^2 - 4x_1x_2",
                "√((x_1 + x_2)^2)",
                "√Δ / |a|",
                "a(1 - x)^2 = b",
            };
            foreach (string lin in samples)
            {
                string mml;
                try { mml = TypeRight.Formulas.MathMLBuilder.Build(lin); }
                catch (Exception ex) { Console.WriteLine("[" + lin + "] BUILD异常 " + ex.Message); continue; }
                Console.WriteLine("[" + lin + "] MathML: " + mml);
                try
                {
                    string omml = TypeRight.Formulas.MathmlInjector.TransformForProbe(mml, xsl);
                    string stats = "sSub=" + Count(omml, "m:sSub") +
                        " sSup=" + Count(omml, "m:sSup") +
                        " sSubSup=" + Count(omml, "m:sSubSup") +
                        " f=" + Count(omml, "m:f") +
                        " rad=" + Count(omml, "m:rad");
                    Console.WriteLine("  OMML " + stats);
                    Console.WriteLine("  片段: " + omml.Substring(0, Math.Min(320, omml.Length)));
                }
                catch (Exception ex) { Console.WriteLine("  XSL异常 " + ex.Message); }
            }
        }

        private static int Count(string s, string tag)
        {
            int n = 0, i = 0;
            while ((i = s.IndexOf(tag, i, StringComparison.Ordinal)) >= 0) { n++; i += tag.Length; }
            return n;
        }
    }
}
