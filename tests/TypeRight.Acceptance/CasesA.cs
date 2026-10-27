using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formatting;
using TypeRight.Profiles;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>TC-01 ~ TC-09：同类选择 / 格式应用与联动 / 撤销与幂等。</summary>
    public static class CasesA
    {
        public static void TC01(TestEnv env)
        {
            env.Begin("TC-01 严格模式（模块样本 → 恰好 7 段）");
            Document doc = env.OpenCopy("tc01");
            try
            {
                Paragraph module = env.FindParagraph(doc, "模块一");
                env.Check(module != null, "找到「模块一：」样本段", null, null, "前缀: 模块一");
                if (module == null) return;

                var sig = FormatSignatureReader.Read(module);
                var hit = SimilarSelector.Match(doc, sig, FormatSignature.IgnoreDims.None);
                string list = string.Join("、", hit.Paragraphs.Take(3).Select(p => Preview(doc, p)));
                env.Check(hit.Count == 7, "严格模式命中数 = 7", "7", hit.Count.ToString(),
                    "样本: 模块一（签名 " + sig + "）前3段: " + list);

                bool containsTitle = hit.Paragraphs.Any(p => ParaText(doc, p).StartsWith("初三数学", StringComparison.Ordinal));
                env.Check(!containsTitle, "命中集不含文档主标题");
            }
            finally { doc.Close(false); }
        }

        public static void TC02(TestEnv env)
        {
            env.Begin("TC-02 宽松模式-忽略字号（→ 恰好 8 段）");
            Document doc = env.OpenCopy("tc02");
            try
            {
                Paragraph module = env.FindParagraph(doc, "模块一");
                var hit = SimilarSelector.Match(doc, module, FormatSignature.IgnoreDims.SizePt);
                env.Check(hit.Count == 8, "忽略字号后命中数 = 8（主标题+7模块；正文样本不捞表格★段）", "8", hit.Count.ToString(),
                    "样本: 模块一");
                bool containsTitle = hit.Paragraphs.Any(p => ParaText(doc, p).StartsWith("初三数学", StringComparison.Ordinal));
                env.Check(containsTitle, "命中集包含文档主标题");
            }
            finally { doc.Close(false); }
        }

        public static void TC03(TestEnv env)
        {
            env.Begin("TC-03 小节标题严格（→ 恰好 16 段）");
            Document doc = env.OpenCopy("tc03");
            try
            {
                Paragraph section = env.FindParagraph(doc, "1.1 ");
                env.Check(section != null, "找到「1.1 」样本段", null, null, "前缀: 1.1 ");
                if (section == null) return;
                var sig = FormatSignatureReader.Read(section);
                var hit = SimilarSelector.Match(doc, sig, FormatSignature.IgnoreDims.None);
                env.Check(hit.Count == 16, "严格模式命中数 = 16", "16", hit.Count.ToString(),
                    "样本: 1.1（签名 " + sig + "）");
            }
            finally { doc.Close(false); }
        }

        public static void TC04(TestEnv env)
        {
            env.Begin("TC-04 正文系三组命中");
            Document doc = env.OpenCopy("tc04");
            try
            {
                Paragraph body = env.FindParagraph(doc, "解一元二次方程的灵魂是");
                env.Check(body != null, "找到正文样本段", null, null, "前缀: 解一元二次方程的灵魂是");
                if (body != null)
                {
                    var sig = FormatSignatureReader.Read(body);
                    var strict = SimilarSelector.Match(doc, body, FormatSignature.IgnoreDims.None);
                    env.Check(strict.Count == 21, "① 严格命中 = 21", "21", strict.Count.ToString(), "样本: 解一元二次方程的灵魂是（" + sig + "）");

                    var loose = SimilarSelector.Match(doc, body, FormatSignature.IgnoreDims.Bold | FormatSignature.IgnoreDims.Color);
                    env.Check(loose.Count == 64, "② 忽略加粗+颜色命中 = 64", "64", loose.Count.ToString(), "同上");
                }

                Paragraph caption = env.FindParagraph(doc, "【图解例图】");
                env.Check(caption != null, "找到图注样本段", null, null, "前缀: 【图解例图】");
                if (caption != null)
                {
                    var sig = FormatSignatureReader.Read(caption);
                    var hit = SimilarSelector.Match(doc, sig, FormatSignature.IgnoreDims.SizePt | FormatSignature.IgnoreDims.Color);
                    env.Check(hit.Count == 7, "③ 图注+忽略字号+颜色命中 = 7（6 图注 + 1 副标题）", "7", hit.Count.ToString(),
                        "样本: 【图解例图】（" + sig + "）");
                }
            }
            finally { doc.Close(false); }
        }

        public static void TC05(TestEnv env)
        {
            env.Begin("TC-05 空命中（字体应用后主标题签名全文唯一）");
            Document doc = env.OpenCopy("tc05");
            try
            {
                Journey.Run(env.App, doc, env, withFormulas: false);
                Paragraph title = env.FindParagraph(doc, "初三数学");
                env.Check(title != null, "找到主标题段", null, null, "前缀: 初三数学");
                if (title == null) return;
                var sig = FormatSignatureReader.Read(title);
                var hit = SimilarSelector.Match(doc, title, FormatSignature.IgnoreDims.None);
                env.Check(hit.Count == 1, "严格命中 = 1（仅当前段落）", "1", hit.Count.ToString(),
                    "样本: 主标题（签名 " + sig + "）");
            }
            finally { doc.Close(false); }
        }

        public static void TC06(TestEnv env)
        {
            env.Begin("TC-06 中西文联动（8 段楷体 + TNR，字号/加粗/颜色保源值）");
            Document doc = env.OpenCopy("tc06");
            try
            {
                Paragraph module = env.FindParagraph(doc, "模块一");
                var hit = SimilarSelector.Match(doc, module, FormatSignature.IgnoreDims.SizePt);
                env.Check(hit.Count == 8, "前置：命中 8 段", "8", hit.Count.ToString(), null);

                // 记录源值（字号/加粗/颜色）
                var before = new List<Tuple<float, int, int>>();
                foreach (Paragraph p in hit.Paragraphs)
                {
                    var f = p.Range.Font;
                    before.Add(Tuple.Create(f.Size, f.Bold, (int)f.Color));
                }

                BatchFormatter.Apply(env.App, doc, hit.Paragraphs,
                    new FormatChanges { EastAsiaFont = "楷体", LatinFont = "Times New Roman" }, "同类格式");

                for (int i = 0; i < hit.Paragraphs.Count; i++)
                {
                    Paragraph p = hit.Paragraphs[i];
                    string anchor = "段: " + Preview(doc, p);
                    var f = p.Range.Font;
                    env.Check(string.Equals(f.NameFarEast, "楷体", StringComparison.Ordinal),
                        "NameFarEast = 楷体", "楷体", f.NameFarEast, anchor);
                    env.Check(string.Equals(f.NameAscii, "Times New Roman", StringComparison.OrdinalIgnoreCase),
                        "NameAscii = Times New Roman", "Times New Roman", f.NameAscii, anchor);
                    env.Check(string.Equals(f.NameOther, "Times New Roman", StringComparison.OrdinalIgnoreCase),
                        "NameOther = Times New Roman", "Times New Roman", f.NameOther, anchor);
                    env.Check(Math.Abs(f.Size - before[i].Item1) < 0.01,
                        "字号保持源值", before[i].Item1.ToString(), f.Size.ToString(), anchor);
                    env.Check(f.Bold == before[i].Item2, "加粗保持源值", before[i].Item2.ToString(), f.Bold.ToString(), anchor);
                    env.Check((int)f.Color == before[i].Item3, "颜色保持源值", before[i].Item3.ToString(), ((int)f.Color).ToString(), anchor);
                }
            }
            finally { doc.Close(false); }
        }

        public static void TC07(TestEnv env)
        {
            env.Begin("TC-07 表格覆盖（11 表内所有 run 宋体 + TNR）");
            Document doc = env.OpenCopy("tc07");
            try
            {
                // 前三轮（模块/小节/正文）后保存，文件级检查表格 run
                Journey.Run(env.App, doc, env, withFormulas: false);
                doc.SaveAs2(env.CopyPath("tc07-saved"));
                doc.Close(false);

                using (var xml = DocXml.Load(env.CopyPath("tc07-saved")))
                {
                    env.Check(xml.TableCount() == 11, "表格数 = 11", "11", xml.TableCount().ToString(), null);

                    int checkedRuns = 0, bad = 0;
                    var firstBad = new List<string>();
                    foreach (XElement tbl in xml.Root.Descendants(DocXml.W + "tbl"))
                        foreach (XElement r in tbl.Descendants(DocXml.W + "r"))
                        {
                            if (DocXml.PlainText(r).Trim().Length == 0) continue;
                            checkedRuns++;
                            XElement rPr = r.Element(DocXml.W + "rPr");
                            string east = null, ascii = null;
                            if (rPr != null)
                            {
                                XElement fonts = rPr.Element(DocXml.W + "rFonts");
                                if (fonts != null)
                                {
                                    east = (string)fonts.Attribute(DocXml.W + "eastAsia");
                                    ascii = (string)fonts.Attribute(DocXml.W + "ascii");
                                }
                            }
                            if (east != "宋体" || ascii != "Times New Roman")
                            {
                                bad++;
                                if (firstBad.Count < 5) firstBad.Add(DocXml.PlainText(r).Substring(0, Math.Min(12, DocXml.PlainText(r).Length)) + "[east=" + east + ",ascii=" + ascii + "]");
                            }
                        }
                    env.Check(checkedRuns > 0, "表格内存在文本 run（检查基数）", ">0", checkedRuns.ToString(), null);
                    env.Check(bad == 0, "表格内全部 run eastAsia=宋体 且 ascii=TNR",
                        "0 处不符", bad + " 处不符：" + string.Join("；", firstBad), "11 张 w:tbl");
                }
            }
            finally { try { doc.Close(false); } catch { } }
        }

        public static void TC08(TestEnv env)
        {
            env.Begin("TC-08 撤销粒度（每步独立撤销单元）");
            Document doc = env.OpenCopy("tc08");
            try
            {
                Paragraph module = env.FindParagraph(doc, "模块一");
                Paragraph section = env.FindParagraph(doc, "1.1 ");
                string moduleEast0 = module.Range.Font.NameAscii;
                string sectionEast0 = section.Range.Font.NameAscii;

                var sig1 = FormatSignatureReader.Read(module);
                var hit1 = SimilarSelector.Match(doc, sig1, FormatSignature.IgnoreDims.SizePt);
                BatchFormatter.Apply(env.App, doc, hit1.Paragraphs, new FormatChanges { EastAsiaFont = "楷体", LatinFont = "Times New Roman" }, "同类格式（模块→楷体）");

                var sig2 = FormatSignatureReader.Read(section);
                var hit2 = SimilarSelector.Match(doc, sig2, FormatSignature.IgnoreDims.None);
                BatchFormatter.Apply(env.App, doc, hit2.Paragraphs, new FormatChanges { EastAsiaFont = "黑体", LatinFont = "Times New Roman" }, "同类格式（小节→黑体）");

                env.Check(module.Range.Font.NameFarEast == "楷体", "撤销前：模块=楷体", "楷体", module.Range.Font.NameFarEast, null);
                env.Check(section.Range.Font.NameFarEast == "黑体", "撤销前：小节=黑体", "黑体", section.Range.Font.NameFarEast, null);

                bool undone = doc.Undo();
                env.Check(undone, "第 1 次 Ctrl+Z 成功");
                env.Check(section.Range.Font.NameAscii == sectionEast0,
                    "撤销后：小节回到源字体（" + sectionEast0 + "）", sectionEast0, section.Range.Font.NameAscii, "段: 1.1 ");
                env.Check(module.Range.Font.NameFarEast == "楷体",
                    "撤销后：模块仍为楷体（仅退最后一步）", "楷体", module.Range.Font.NameFarEast, "段: 模块一");

                bool undone2 = doc.Undo();
                env.Check(undone2, "第 2 次 Ctrl+Z 成功");
                env.Check(module.Range.Font.NameAscii == moduleEast0,
                    "再撤销后：模块回到源字体（" + moduleEast0 + "）", moduleEast0, module.Range.Font.NameAscii, "段: 模块一");
            }
            finally { doc.Close(false); }
        }

        public static void TC09(TestEnv env)
        {
            env.Begin("TC-09 幂等（重复应用相同字体 0 改动）");
            Document doc = env.OpenCopy("tc09");
            try
            {
                Paragraph module = env.FindParagraph(doc, "模块一");
                var hit = SimilarSelector.Match(doc, module, FormatSignature.IgnoreDims.SizePt);
                var changes = new FormatChanges { EastAsiaFont = "楷体", LatinFont = "Times New Roman" };

                var stats1 = BatchFormatter.Apply(env.App, doc, hit.Paragraphs, changes, "同类格式");
                var stats2 = BatchFormatter.Apply(env.App, doc, hit.Paragraphs, changes, "同类格式(重复)");
                env.Check(stats1.Changed > 0, "首次应用发生改动", ">0", stats1.Changed.ToString(), null);
                env.Check(stats2.Changed == 0, "二次应用 0 改动", "0", stats2.Changed.ToString(), "样本: 模块一（忽略字号，8 段）");
            }
            finally { doc.Close(false); }
        }

        // ---- 辅助 ----

        internal static string ParaText(Document doc, Paragraph p)
        {
            try { return (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v'); }
            catch { return ""; }
        }

        internal static string Preview(Document doc, Paragraph p)
        {
            string t = ParaText(doc, p);
            return t.Length > 12 ? t.Substring(0, 12) : t;
        }
    }
}
