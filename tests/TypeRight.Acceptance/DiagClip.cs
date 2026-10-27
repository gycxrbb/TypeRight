using System;
using NetOffice.WordApi;

namespace TypeRight.Acceptance
{
    /// <summary>剪贴板 MathML → Word 粘贴注入探针（TC-19 管线关键路径）。</summary>
    internal static class DiagClip
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            string mathml =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                "<mfrac>" +
                "<mrow><mo>-</mo><mi>b</mi><mo>&#x00B1;</mo><msqrt><mi mathvariant=\"normal\">&#x0394;</mi></msqrt></mrow>" +
                "<mrow><mn>2</mn><mi>a</mi></mrow>" +
                "</mfrac>" +
                "</math>";

            string mathml2 =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                "<msubsup><mi>x</mi><mn>1</mn><mn>2</mn></msubsup>" +
                "<mo>+</mo>" +
                "<mfrac><mi>a</mi><mi>b</mi></mfrac>" +
                "</math>";

            // TC-19 失败串复测（INJECT 回退清单）
            string[] failing =
            {
                "x_2", "k - 1 ≠ 0", "(k-1)x^2", "x", "y_1 < y_2", "x = h",
                "= (3x_1 - 1) + 2(3 - x_1) = x_1 + 5", "x_1^2 + 2x_2",
                "k - 1 ≠ 0 ⟹ k ≠ 1", "(k - 1)x^2 + 2kx + k + 2 = 0", "p ≥ 0", "(-b±√(b^2-4ac))/(2a)"
            };
            foreach (string lin in failing)
            {
                string mml;
                try { mml = TypeRight.Formulas.MathMLBuilder.Build(lin); }
                catch (Exception ex) { Console.WriteLine("  BUILD失败[" + lin + "]: " + ex.Message); continue; }
                TryPaste(env, "F:" + lin, mml, r => System.Windows.Forms.Clipboard.SetText(r, System.Windows.Forms.TextDataFormat.UnicodeText));
            }
        }

        private static void TryPaste(TestEnv env, string name, string mathml, Action<string> setClipboard)
        {
            Document doc = env.App.Documents.Add();
            try
            {
                var rng = doc.Range(0, 0);
                rng.InsertAfter("XX" + name + "YY");
                var find = doc.Content;
                find.Find.Text = name;
                find.Find.Forward = true;
                find.Find.Wrap = NetOffice.WordApi.Enums.WdFindWrap.wdFindStop;
                if (!find.Find.Execute()) { Console.WriteLine("  " + name + ": 定位失败"); doc.Close(false); return; }

                // 缩到标记词本身（不含 XX/YY）
                var target = doc.Range(find.Start, find.End);
                try
                {
                    setClipboard(mathml);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  " + name + ": 剪贴板写入失败 " + ex.Message);
                    doc.Close(false);
                    return;
                }

                target.Select();
                int zonesBefore = doc.OMaths.Count;
                try
                {
                    env.App.Selection.Paste();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  " + name + ": 粘贴异常 " + ex.Message);
                    doc.Close(false);
                    return;
                }

                int zones = doc.OMaths.Count;
                string info = zones > zonesBefore ? "oMath+" + (zones - zonesBefore) : "无新公式区";
                Console.WriteLine("  " + name + ": 粘贴OK " + info + " | 文档oMath=" + zones);

                doc.SaveAs2(env.CopyPath("clip-" + name.Substring(0, 2)));
                doc.Close(false);
                using (var xml = DocXml.Load(env.CopyPath("clip-" + name.Substring(0, 2))))
                {
                    Console.WriteLine("    文件级: oMath=" + xml.OMathCount() +
                        " m:f=" + xml.MathElementCount("f") +
                        " m:rad=" + xml.MathElementCount("rad") +
                        " sSubSup=" + xml.MathElementCount("sSubSup") +
                        " | 剩余文本含XX: " + DocXml.PlainText(xml.Root).Contains("XX"));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("  " + name + ": 外层异常 " + ex.Message);
                try { doc.Close(false); } catch { }
            }
        }
    }
}
