using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using NetOffice.WordApi;
using TypeRight.Formatting;
using TypeRight.Profiles;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>TC-19（分式内部结构）/ TC-20（表格字号保持）。</summary>
    public static class CasesD
    {
        // ---- TC-19：MML2OMML 管线后的结构完整性 ----

        public static void TC19(TestEnv env)
        {
            env.Begin("TC-19 分式内部结构完整（MML2OMML 管线）");
            string actual = env.CopyPath("tc19");
            Document doc = env.OpenCopy("tc19");
            try
            {
                var journey = Journey.Run(env.App, doc, env, withFormulas: true);
                Console.WriteLine("  公式: " + (journey.Formula != null ? journey.Formula.SummaryText() : "-"));
                doc.SaveAs2(actual);
                doc.Close(false);
            }
            catch { try { doc.Close(false); } catch { } }

            using (var xml = DocXml.Load(actual))
            {
                // ① 上下标结构总数 = 201 ±2
                int sub = xml.MathElementCount("sSub");
                int sup = xml.MathElementCount("sSup");
                int subsup = xml.MathElementCount("sSubSup");
                int total = sub + sup + subsup;
                env.Check(Math.Abs(total - 201) <= 2, "上下标结构总数 = 201（±2）",
                    "199~203", total.ToString(), "sSub=" + sub + " sSup=" + sup + " sSubSup=" + subsup);

                // ② Unicode 上下标残留 = 0
                int residue = xml.UnicodeMathCharsInPlainText();
                env.Check(residue == 0, "w:t 中 Unicode 上下标/根号字符 = 0", "0", residue.ToString(), null);

                // ③ 弦长秒杀公式：m:f 的 num 内含 m:rad，radicand 内含 m:sSup
                var p = xml.FindParagraphByPrefix("【弦长秒杀公式】");
                if (!CheckFound(env, p != null, "找到「【弦长秒杀公式】」段", "前缀: 【弦长秒杀公式】")) return;

                bool ok = false;
                string detail = "未找到匹配结构";
                foreach (var f in OmmlAssert.Fractions(p))
                {
                    var num = OmmlAssert.Num(f);
                    if (num == null) continue;
                    foreach (var rad in num.Descendants(DocXml.M + "rad"))
                    {
                        if (rad.Descendants(DocXml.M + "sSup").Any() || rad.Descendants(DocXml.M + "sSubSup").Any())
                        {
                            ok = true;
                            detail = "m:f/num/m:rad(" + DocXml.MathText(rad).Substring(0, Math.Min(10, DocXml.MathText(rad).Length)) + ") 含 sSup";
                            break;
                        }
                    }
                    if (ok) break;
                }
                env.Check(ok, "「【弦长秒杀公式】」m:f 的 num 内含 m:rad 且 radicand 内含 m:sSup", "结构存在", detail, "【弦长秒杀公式】");

                // 补充：m:f 总量参考（黄金 25）
                int mf = xml.MathElementCount("f");
                Console.WriteLine("  [参考] m:f=" + mf + "（黄金 25，±2 为 23~27）");
            }
        }

        private static bool CheckFound(TestEnv env, bool cond, string what, string anchor)
        {
            env.Check(cond, what, "存在", cond ? "存在" : "未找到", anchor);
            return cond;
        }

        // ---- TC-20：表格内非公式 run 的 sz/b/color 逐一相等 ----

        public static void TC20(TestEnv env)
        {
            env.Begin("TC-20 表格内字号不丢失（rPr 深拷贝）");
            Document doc = env.OpenCopy("tc20");
            try
            {
                var stats = new FormulaConverter(doc, null, TypeRight.Formatting.FormulaConverter.ResolveMml2OmmlPath(env.App)).ConvertDocument();
                Console.WriteLine("  " + stats.SummaryText());
                doc.SaveAs2(env.CopyPath("tc20-out"));
                doc.Close(false);
            }
            catch { try { doc.Close(false); } catch { } }

            var before = ExtractTableRunProps(env.OriginalDocx);
            var after = ExtractTableRunProps(env.CopyPath("tc20-out"));

            env.Check(before.Count > 0 && after.Count > 0, "表格 run 提取基数 > 0",
                ">0", before.Count + "/" + after.Count, null);
            env.Check(before.Count == after.Count, "非公式 run 数量不变（公式 run 剔除后）",
                before.Count.ToString(), after.Count.ToString(), null);

            if (before.Count != after.Count) return;

            int mismatch = 0;
            var samples = new List<string>();
            for (int i = 0; i < before.Count; i++)
            {
                var b = before[i];
                var a = after[i];
                if (b.Text != a.Text)
                {
                    mismatch++;
                    if (samples.Count < 3) samples.Add("文本漂移[" + b.Text + "→" + a.Text + "]");
                    continue;
                }
                if (!SzEq(b, a) || b.Bold != a.Bold || !ColorEq(b, a))
                {
                    mismatch++;
                    if (samples.Count < 5)
                        samples.Add("'" + (b.Text.Length > 6 ? b.Text.Substring(0, 6) : b.Text) + "' sz:" + SzStr(b) + "→" + SzStr(a) +
                            " b:" + b.Bold + "→" + a.Bold + " c:" + b.Color + "→" + a.Color);
                }
            }
            env.Check(mismatch == 0, "表格内所有非公式 run 的 w:sz/w:b/w:color 逐一相等",
                "0 处不一致", mismatch + " 处：" + string.Join("；", samples), "11 张 w:tbl");
        }

        private sealed class RunProps
        {
            public string Text;
            public string Sz;      // null = 无显式
            public bool Bold;
            public string Color;
        }

        private static bool SzEq(RunProps b, RunProps a)
        {
            if (b.Sz == null && a.Sz == null) return true;
            if (b.Sz == null || a.Sz == null) return false;
            return b.Sz == a.Sz;
        }

        private static string SzStr(RunProps r) { return r.Sz ?? "(无)"; }

        private static bool ColorEq(RunProps b, RunProps a)
        {
            return string.Equals(b.Color ?? "auto", a.Color ?? "auto", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>提取全部表格内非公式（不在 m:oMath 内）文本 run 的 (text, sz, bold, color)。</summary>
        private static List<RunProps> ExtractTableRunProps(string docx)
        {
            var result = new List<RunProps>();
            using (var xml = DocXml.Load(docx))
            {
                foreach (XElement tbl in xml.Root.Descendants(DocXml.W + "tbl"))
                {
                    foreach (XElement r in tbl.Descendants(DocXml.W + "r"))
                    {
                        if (r.Ancestors(DocXml.M + "oMath").Any()) continue;   // 公式 run 剔除
                        string text = DocXml.PlainText(r);
                        if (text.Trim().Length == 0) continue;
                        var rp = new RunProps { Text = text };
                        XElement rPr = r.Element(DocXml.W + "rPr");
                        if (rPr != null)
                        {
                            XElement sz = rPr.Element(DocXml.W + "sz");
                            if (sz != null) rp.Sz = (string)sz.Attribute(DocXml.W + "val");
                            XElement b = rPr.Element(DocXml.W + "b");
                            if (b != null)
                            {
                                string v = (string)b.Attribute(DocXml.W + "val");
                                rp.Bold = v != "0" && v != "false";
                            }
                            XElement c = rPr.Element(DocXml.W + "color");
                            if (c != null) rp.Color = (string)c.Attribute(DocXml.W + "val");
                        }
                        result.Add(rp);
                    }
                }
            }
            return result;
        }
    }
}
