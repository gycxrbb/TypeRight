using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formatting;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>TC-10 ~ TC-14：公式转换 / 黄金旅程 / 内容零损失 / 源格式保持。</summary>
    public static class CasesB
    {
        // ---- §4.1 总量断言（±2 容差，超出输出差异清单）----

        private static void CheckTotals(TestEnv env, DocXml xml, string label)
        {
            int oMath = xml.OMathCount();
            env.Check(Math.Abs(oMath - 255) <= 2, label + "：m:oMath = 255（±2）", "253~257", oMath.ToString(), "document.xml 全文");

            int mf = xml.MathElementCount("f");
            env.Check(Math.Abs(mf - 25) <= 2, label + "：m:f（分式）= 25（±2）", "23~27", mf.ToString(), "document.xml 全文");

            int rad = xml.MathElementCount("rad");
            env.Check(Math.Abs(rad - 7) <= 2, label + "：m:rad（根式）= 7（±2）", "5~9", rad.ToString(), "document.xml 全文");

            int struct3 = xml.MathElementCount("sSub") + xml.MathElementCount("sSup") + xml.MathElementCount("sSubSup");
            env.Check(Math.Abs(struct3 - 201) <= 2, label + "：上下标结构（sSub+sSup+sSubSup）= 201（±2）", "199~203", struct3.ToString(),
                "sSub=" + xml.MathElementCount("sSub") + " sSup=" + xml.MathElementCount("sSup") + " sSubSup=" + xml.MathElementCount("sSubSup"));

            int unicode = xml.UnicodeMathCharsInPlainText();
            env.Check(unicode == 0, label + "：w:t 中 Unicode 上下标/根号字符 = 0", "0", unicode.ToString(), "⁰¹²³⁴⁵⁶⁷⁸⁹ⁿ₀₁₂₃₄√ 全集");
        }

        // ---- §4.3 结构锚点 ----

        private static void CheckStructuralAnchors(TestEnv env, DocXml xml, string label)
        {
            // ① 不等实根：m:f，num 含 m:rad(Δ) 与 ±，den 文本 2a
            var p1 = xml.FindParagraphByPrefix("① 不等实根");
            bool has1 = p1 != null; env.Check(has1, label + "：找到锚点段「① 不等实根」", "段落存在", "未找到", "前缀: ① 不等实根");
            if (has1)
            {
                var f = OmmlAssert.Fractions(p1).FirstOrDefault(x =>
                    x.Descendants(DocXml.M + "rad").Any() &&
                    DocXml.MathText(x).Contains("±") &&
                    OmmlAssert.DenTextIs(x, "2a"));
                env.Check(f != null, label + "：「① 不等实根」含 m:f(num=m:rad Δ+±, den=2a)", "结构存在", f == null ? "未找到该结构" : "存在", "① 不等实根");
            }

            // 【公式核心】：m:sSub 表达 x₁/x₂；两个 m:f den=a
            var p2 = xml.FindParagraphByPrefix("【公式核心】");
            bool has2 = p2 != null; env.Check(has2, label + "：找到锚点段「【公式核心】」", "段落存在", "未找到", "前缀: 【公式核心】");
            if (has2)
            {
                int subCount = p2.Descendants(DocXml.M + "sSub").Count();
                env.Check(subCount >= 2, label + "：「【公式核心】」sSub ≥ 2（x₁、x₂）", "≥2", subCount.ToString(), "【公式核心】");
                int denA = OmmlAssert.Fractions(p2).Count(x => OmmlAssert.DenTextIs(x, "a"));
                env.Check(denA >= 2, label + "：「【公式核心】」den=a 的 m:f ≥ 2", "≥2", denA.ToString(), "【公式核心】");
            }

            // 2. 差方公式：sSubSup(x,1,2)；m:f num=m:rad(Δ)
            var p3 = xml.FindParagraphByPrefix("2. 差方公式");
            bool has3 = p3 != null; env.Check(has3, label + "：找到锚点段「2. 差方公式」", "段落存在", "未找到", "前缀: 2. 差方公式");
            if (has3)
            {
                bool subSup = p3.Descendants(DocXml.M + "sSubSup").Any(x =>
                    DocXml.MathText(x.Element(DocXml.M + "e")).Trim() == "x" &&
                    DocXml.MathText(x.Element(DocXml.M + "sub")).Trim() == "1" &&
                    DocXml.MathText(x.Element(DocXml.M + "sup")).Trim() == "2");
                env.Check(subSup, label + "：「2. 差方公式」sSubSup(x,1,2)", "存在", "未找到", "2. 差方公式");

                bool fRad = OmmlAssert.Fractions(p3).Any(x => OmmlAssert.Num(x) != null &&
                    OmmlAssert.Num(x).Descendants(DocXml.M + "rad").Any(rad => DocXml.MathText(rad).Contains("Δ")));
                env.Check(fRad, label + "：「2. 差方公式」m:f(num=m:rad Δ)", "存在", "未找到", "2. 差方公式");
            }

            // 【弦长秒杀公式】：m:rad 内含 m:sSup；m:f(√Δ/|a|)
            var p4 = xml.FindParagraphByPrefix("【弦长秒杀公式】");
            bool has4 = p4 != null; env.Check(has4, label + "：找到锚点段「【弦长秒杀公式】」", "段落存在", "未找到", "前缀: 【弦长秒杀公式】");
            if (has4)
            {
                bool radSup = p4.Descendants(DocXml.M + "rad").Any(rad => rad.Descendants(DocXml.M + "sSup").Any());
                env.Check(radSup, label + "：「【弦长秒杀公式】」m:rad 内含 m:sSup", "存在", "未找到", "【弦长秒杀公式】");
                bool fRad2 = OmmlAssert.Fractions(p4).Any(x => OmmlAssert.Num(x) != null &&
                    OmmlAssert.Num(x).Descendants(DocXml.M + "rad").Any());
                env.Check(fRad2, label + "：「【弦长秒杀公式】」m:f(√Δ/|a|)", "存在", "未找到", "【弦长秒杀公式】");
            }

            // 表格「公式法」行（含 当 Δ ≥ 0）：单元格内 m:f，num 含 m:rad(Δ)
            var cell = xml.FindTableCellContaining("当 Δ ≥ 0");
            bool has5 = cell != null; env.Check(has5, label + "：找到表格行（当 Δ ≥ 0）", "单元格存在", "未找到", "w:tc 文本包含: 当 Δ ≥ 0");
            if (has5)
            {
                bool fRadCell = OmmlAssert.Fractions(cell).Any(x => OmmlAssert.Num(x) != null &&
                    OmmlAssert.Num(x).Descendants(DocXml.M + "rad").Any(rad => DocXml.MathText(rad).Contains("Δ")));
                env.Check(fRadCell, label + "：表格行 m:f(num=m:rad Δ)", "存在", "未找到", "w:tc: 当 Δ ≥ 0");
            }
        }

        // ---- §4.4 负向断言 ----

        private static void CheckNegatives(TestEnv env, DocXml xml, string label)
        {
            string plain = DocXml.PlainText(xml.Root);
            string[] mustStayText =
            {
                "1.1", "6.5", "70%", "90%", "15 秒", "vs", "➔", "①", "②", "③"
            };
            foreach (string s in mustStayText)
            {
                env.Check(plain.Contains(s), label + "：保持纯文本「" + s + "」", "w:t 中存在", "缺失（可能被误转公式）", "全文 w:t 拼接");
            }
        }

        // ---- §4.2 字体断言（按前缀锚点 + 计数）----

        public static void CheckFontAnchors(TestEnv env, DocXml xml, string label)
        {
            CheckFontAnchor(env, xml, label, "初三数学代数终极大统领", "楷体");
            CheckFontAnchor(env, xml, label, "模块", "楷体");
            CheckFontAnchor(env, xml, label, "1.1 ", "黑体");

            env.Check(xml.CountParagraphsByEastAsia("楷体") == 8, label + "：楷体段 = 8", "8", xml.CountParagraphsByEastAsia("楷体").ToString(), "首 run w:rFonts/@eastAsia");
            env.Check(xml.CountParagraphsByEastAsia("黑体") == 16, label + "：黑体段 = 16", "16", xml.CountParagraphsByEastAsia("黑体").ToString(), "首 run w:rFonts/@eastAsia");
            env.Check(xml.CountParagraphsByEastAsia("宋体") == 119, label + "：宋体段 = 119（黄金实测；规格原文 71 为正文口径：64+7）", "119", xml.CountParagraphsByEastAsia("宋体").ToString(), "首 run w:rFonts/@eastAsia");
        }

        private static void CheckFontAnchor(TestEnv env, DocXml xml, string label, string prefix, string eastAsia)
        {
            foreach (XElement p in xml.FindParagraphsByPrefix(prefix))
            {
                var rf = DocXml.FirstRunFonts(p);
                string anchor = "段: " + DocXml.PlainText(p).Substring(0, Math.Min(12, DocXml.PlainText(p).Length));
                env.Check(rf != null && rf.EastAsia == eastAsia, label + "：中文字体 = " + eastAsia, eastAsia, rf == null ? "(无 run)" : rf.EastAsia, anchor);
                env.Check(rf != null && rf.Ascii == "Times New Roman", label + "：西文字体 = Times New Roman", "Times New Roman", rf == null ? "(无 run)" : rf.Ascii, anchor);
            }
        }

        // ---- 用例 ----

        public static void TC10(TestEnv env)
        {
            env.Begin("TC-10 公式转换总量与结构");
            Document doc = env.OpenCopy("tc10");
            try
            {
                Journey.Run(env.App, doc, env, withFormulas: true);
                doc.SaveAs2(env.CopyPath("tc10-out"));
                doc.Close(false);

                using (var xml = DocXml.Load(env.CopyPath("tc10-out")))
                {
                    CheckTotals(env, xml, "TC-10");
                    CheckStructuralAnchors(env, xml, "TC-10");
                    CheckNegatives(env, xml, "TC-10");
                }
            }
            finally { try { doc.Close(false); } catch { } }
        }

        public static void TC11(TestEnv env)
        {
            env.Begin("TC-11 失败安全（括号不闭合片段原样保留）");
            Document doc = env.OpenCopy("tc11");
            try
            {
                // 注入坏片段段落
                var range = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                range.InsertAfter("验证片段 y = a(x - h 说明文字");
                range.InsertParagraphAfter();

                var stats = new FormulaConverter(doc, null, null).ConvertDocument();
                Console.WriteLine("  " + stats.SummaryText());

                env.Check(stats.Failed >= 1, "失败计数 ≥ 1", "≥1", stats.Failed.ToString(), "注入片段: y = a(x - h");
                env.Check(stats.Converted > 100, "其余公式照常转换（>100）", ">100", stats.Converted.ToString(), null);

                string text = doc.Content.Text ?? "";
                env.Check(text.Contains("y = a(x - h"), "坏片段原样保留为文本", "包含 'y = a(x - h'", "已缺失（被改动）", "注入段");
                // 坏片段未被包进公式对象：该段落不含 oMath 覆盖它 —— 用查找定位粗验
                var find = doc.Content;
                find.Find.ClearFormatting();
                find.Find.Text = "y = a(x - h";
                find.Find.Forward = true;
                find.Find.Wrap = WdFindWrap.wdFindStop;
                bool found = find.Find.Execute();
                env.Check(found, "坏片段可通过纯文本查找命中（未进公式对象）", "可找到", found ? "可找到" : "找不到", "Find: y = a(x - h");
                if (found) env.Check(find.OMaths.Count == 0, "坏片段区间内 oMath = 0", "0", find.OMaths.Count.ToString(), "Find 结果区间");
            }
            finally { doc.Close(false); }
        }

        public static void TC12(TestEnv env)
        {
            env.Begin("TC-12 黄金旅程（4 轮同类选择 + 公式转换 → §4.1~4.4）");
            Document doc = env.OpenCopy("actual");
            try
            {
                var journey = Journey.Run(env.App, doc, env, withFormulas: true);
                Console.WriteLine(string.Format("  轮次命中: {0}/{1}/{2}/{3}  公式: {4}  总耗时 {5}ms",
                    journey.Round1Hits, journey.Round2Hits, journey.Round3Hits, journey.Round4Hits,
                    journey.Formula != null ? journey.Formula.Converted + " 转换 / " + journey.Formula.Failed + " 失败" : "-",
                    journey.TotalMs));

                doc.SaveAs2(env.CopyPath("actual"));
                doc.Close(false);

                using (var xml = DocXml.Load(env.CopyPath("actual")))
                {
                    CheckTotals(env, xml, "TC-12");
                    CheckStructuralAnchors(env, xml, "TC-12");
                    CheckNegatives(env, xml, "TC-12");
                    CheckFontAnchors(env, xml, "TC-12");
                }
            }
            finally { try { doc.Close(false); } catch { } }
        }

        public static void TC13(TestEnv env)
        {
            env.Begin("TC-13 内容零损失（段落/表格/图片数 + 中文字符流）");
            string actual = env.CopyPath("actual");
            if (!System.IO.File.Exists(actual))
            {
                env.Check(false, "前置：actual.docx 存在（先跑 TC-12）", "存在", "缺失", "tests/output/actual.docx");
                return;
            }
            using (var a = DocXml.Load(actual))
            using (var o = DocXml.Load(env.OriginalDocx))
            {
                env.Check(o.NonEmptyParagraphCount() == 145, "原始非空段落 = 145（含表格段落；规格原文 95 为正文口径）", "145", o.NonEmptyParagraphCount().ToString(), null);
                env.Check(a.NonEmptyParagraphCount() == 145, "产出非空段落 = 145", "145", a.NonEmptyParagraphCount().ToString(), null);
                env.Check(a.TableCount() == 11 && o.TableCount() == 11, "表格数 = 11", "11", a.TableCount() + "/" + o.TableCount(), null);
                env.Check(a.ImageCount() == 6 && o.ImageCount() == 6, "图片数 = 6", "6", a.ImageCount() + "/" + o.ImageCount(), null);

                string streamA = a.CjkStream();
                string streamO = o.CjkStream();
                env.Check(streamA == streamO, "全文中文字符流逐字相等",
                    streamO.Length + " 个 CJK 字符", streamA == streamO ? "相等" : "长度 " + streamA.Length + "（原始 " + streamO.Length + "）",
                    "w:t 顺序拼接过滤 CJK");
                if (streamA != streamO)
                {
                    int i = 0;
                    while (i < Math.Min(streamA.Length, streamO.Length) && streamA[i] == streamO[i]) i++;
                    Console.WriteLine("         首个差异位置 " + i + "：期望『" + SafeSub(streamO, i) + "』实际『" + SafeSub(streamA, i) + "』");
                }
            }
        }

        private static string SafeSub(string s, int i)
        {
            if (s == null || i >= s.Length) return "(越界)";
            int len = Math.Min(10, s.Length - i);
            return s.Substring(i, len);
        }

        public static void TC14(TestEnv env)
        {
            env.Begin("TC-14 源格式保持（字号/加粗/颜色与原始逐段一致）");
            string actual = env.CopyPath("actual");
            if (!System.IO.File.Exists(actual))
            {
                env.Check(false, "前置：actual.docx 存在（先跑 TC-12）", "存在", "缺失", "tests/output/actual.docx");
                return;
            }
            using (var a = DocXml.Load(actual))
            using (var o = DocXml.Load(env.OriginalDocx))
            {
                // 按顺序配对非空段（内容零损失已保证可配对），对比首 run 的字号/加粗/颜色
                var listO = o.Paragraphs().Where(p => (DocXml.PlainText(p) + DocXml.MathText(p)).Trim().Length > 0).ToList();
                var listA = a.Paragraphs().Where(p => (DocXml.PlainText(p) + DocXml.MathText(p)).Trim().Length > 0).ToList();
                env.Check(listO.Count == listA.Count, "段落配对基数一致", listO.Count.ToString(), listA.Count.ToString(), null);

                int mismatch = 0, bold43Checked = 0;
                var firstBad = new List<string>();
                int n = Math.Min(listO.Count, listA.Count);
                for (int i = 0; i < n; i++)
                {
                    var fo = DocXml.FirstRunFonts(listO[i]);
                    var fa = DocXml.FirstRunFonts(listA[i]);
                    if (fo == null || fa == null) continue;

                    bool sizeEq = NullableEq(fo.SizePt, fa.SizePt);
                    bool boldEq = NullableEqBool(fo.Bold, fa.Bold);
                    bool colorEq = string.Equals(fo.Color, fa.Color, StringComparison.OrdinalIgnoreCase);

                    // 原始 1E1E1E 加粗要点段：专项核对（43 处）
                    if (fo.Color == "1E1E1E" && fo.Bold == true)
                    {
                        bold43Checked++;
                        if (!(boldEq && colorEq && sizeEq)) mismatch++;
                    }
                    else if (!(sizeEq && boldEq && colorEq))
                    {
                        mismatch++;
                        if (firstBad.Count < 5)
                            firstBad.Add("段『" + DocXml.PlainText(listO[i]).Substring(0, Math.Min(10, DocXml.PlainText(listO[i]).Length)) +
                                "』期望 sz=" + fo.SizePt + "/b=" + fo.Bold + "/c=" + fo.Color +
                                " 实际 sz=" + fa.SizePt + "/b=" + fa.Bold + "/c=" + fa.Color);
                    }
                }
                env.Check(bold43Checked == 43, "1E1E1E 加粗要点段识别数 = 43", "43", bold43Checked.ToString(), "原始文件首 run color=1E1E1E 且加粗");
                env.Check(mismatch == 0, "全部锚点段字号/加粗/颜色与原始一致",
                    "0 处不一致（含 43 要点段）", mismatch + " 处不一致：" + string.Join("；", firstBad), "顺序配对非空段落");
            }
        }

        private static bool NullableEq(double? a, double? b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;
            return Math.Abs(a.Value - b.Value) < 0.01;
        }

        private static bool NullableEqBool(bool? a, bool? b)
        {
            return a == b; // null==null / true==true / false==false
        }
    }
}
