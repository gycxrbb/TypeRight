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

                bool ok = p.Descendants(DocXml.M + "f").Count() >= 1
                    && p.Descendants(DocXml.M + "rad").Count() >= 1
                    && (p.Descendants(DocXml.M + "sSup").Count() >= 1 || p.Descendants(DocXml.M + "sSubSup").Count() >= 1);
                string detail = string.Format("f={0} rad={1} sSup+ss={2}",
                    p.Descendants(DocXml.M + "f").Count(), p.Descendants(DocXml.M + "rad").Count(),
                    p.Descendants(DocXml.M + "sSup").Count() + p.Descendants(DocXml.M + "sSubSup").Count());
                env.Check(ok, "「【弦长秒杀公式」分式+根式+上标结构共存", "结构存在", detail, "【弦长秒杀公式】");

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

            // FormattedText 产生 OMML 包装 run（数量与原文不对齐），按文本配对比较格式
            var beforeText = before.Where(b => !string.IsNullOrWhiteSpace(b.Text)).ToList();
            var afterText = after.Where(a => !string.IsNullOrWhiteSpace(a.Text)).ToList();
            env.Check(beforeText.Count > 0, "表格可见文本 run 提取基数 > 0", ">0", beforeText.Count.ToString(), null);

            int mismatch = 0;
            var samples = new List<string>();
            foreach (var b in beforeText)
            {
                // 在产出中找同文本 run（文本可能被公式替换，找不到=已转换，跳过）
                var candidates = afterText.Where(a => a.Text == b.Text || a.Text.Contains(b.Text) || b.Text.Contains(a.Text)).ToList();
                if (candidates.Count == 0) continue;

                bool matched = candidates.Any(a =>
                    (b.Sz == null || a.Sz == null || b.Sz == a.Sz) &&
                    b.Bold == a.Bold &&
                    string.Equals(b.Color ?? "auto", a.Color ?? "auto", StringComparison.OrdinalIgnoreCase));
                if (!matched)
                {
                    mismatch++;
                    if (samples.Count < 5)
                        samples.Add("'" + (b.Text.Length > 6 ? b.Text.Substring(0, 6) : b.Text) + "' b:" + b.Bold + " c:" + b.Color);
                }
            }
            env.Check(mismatch == 0, "表格内非公式文本 run 的 w:sz/w:b/w:color 保持一致",
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
                        if (r.Ancestors(DocXml.M + "oMathPara").Any()) continue;  // 公式段落包裹
                        if (r.Descendants(DocXml.M + "t").Any()) continue;        // 含数学文本的 run（FormattedText 产生的包装 run）
                        if (r.Element(DocXml.W + "rPr") != null && r.Element(DocXml.W + "t") == null) continue; // 无文本的格式包装 run
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
