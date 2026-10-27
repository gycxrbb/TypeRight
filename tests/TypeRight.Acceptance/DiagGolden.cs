using System;
using System.Linq;
using System.Text;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>黄金文件取证：字体普查 / ★段状态 / 公式结构总量 / 与规格数字对照。</summary>
    internal static class DiagGolden
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("══ 黄金文件取证: " + System.IO.Path.GetFileName(env.GoldenDocx) + " ══");
            using (var xml = DocXml.Load(env.GoldenDocx))
            {
                Console.WriteLine("非空段(w:t+m:t): " + NonEmptyInclMath(xml) +
                    " | w:t-only: " + xml.NonEmptyParagraphCount() +
                    " | 表格: " + xml.TableCount() + " | 图片: " + xml.ImageCount());
                Console.WriteLine("oMath=" + xml.OMathCount() +
                    " m:f=" + xml.MathElementCount("f") +
                    " m:rad=" + xml.MathElementCount("rad") +
                    " sSub=" + xml.MathElementCount("sSub") +
                    " sSup=" + xml.MathElementCount("sSup") +
                    " sSubSup=" + xml.MathElementCount("sSubSup") +
                    " | w:t中Unicode上下标字符=" + xml.UnicodeMathCharsInPlainText());
                Console.WriteLine("字体段数: 楷体=" + xml.CountParagraphsByEastAsia("楷体") +
                    " 黑体=" + xml.CountParagraphsByEastAsia("黑体") +
                    " 宋体=" + xml.CountParagraphsByEastAsia("宋体") +
                    " YaHei=" + xml.CountParagraphsByEastAsia("微软雅黑") +
                    " +中文正文=" + xml.CountParagraphsByEastAsia("+中文正文"));

                Console.WriteLine();
                Console.WriteLine("── 关键锚点段（黄金）──");
                Dump(xml, "初三数学代数终极大统领");
                foreach (var p in xml.FindParagraphsByPrefix("模块").Take(3)) DumpEl(xml, p, "模块");
                Dump(xml, "1.1 ");
                Dump(xml, "——涵盖");
                foreach (var p in xml.FindParagraphsByPrefix("【图解例图】").Take(2)) DumpEl(xml, p, "图注");

                Console.WriteLine();
                Console.WriteLine("── ★ 段（黄金中 8 个）──");
                foreach (var p in xml.FindParagraphsByPrefix("★").Take(3)) DumpEl(xml, p, "★");

                Console.WriteLine();
                Console.WriteLine("── 表格 run 字体抽样（黄金）──");
                int kai = 0, song = 0, hei = 0, other = 0, none = 0;
                string sampleOther = null;
                foreach (var tbl in xml.Root.Descendants(DocXml.W + "tbl"))
                    foreach (var r in tbl.Descendants(DocXml.W + "r"))
                    {
                        if (DocXml.PlainText(r).Trim().Length == 0) continue;
                        var rPr = r.Element(DocXml.W + "rPr");
                        string east = null;
                        if (rPr != null)
                        {
                            var f = rPr.Element(DocXml.W + "rFonts");
                            if (f != null) east = (string)f.Attribute(DocXml.W + "eastAsia");
                        }
                        if (east == null) none++;
                        else if (east == "楷体") kai++;
                        else if (east == "宋体") song++;
                        else if (east == "黑体") hei++;
                        else { other++; if (sampleOther == null) sampleOther = east + "『" + DocXml.PlainText(r).Substring(0, Math.Min(10, DocXml.PlainText(r).Length)) + "』"; }
                    }
                Console.WriteLine("  表格 run: 宋体=" + song + " 楷体=" + kai + " 黑体=" + hei + " 无eastAsia=" + none + " 其他=" + other + (sampleOther != null ? " 如" + sampleOther : ""));

                Console.WriteLine();
                Console.WriteLine("── 黄金公式区清单（255 个的线性文本）──");
                {
                    int zi = 0;
                    var zoneTexts = new System.Collections.Generic.List<string>();
                    foreach (var om in xml.Root.Descendants(DocXml.M + "oMath"))
                    {
                        zi++;
                        zoneTexts.Add(DocXml.MathText(om));
                    }
                    foreach (string z in zoneTexts.Take(60))
                        Console.WriteLine("  zone[" + z + "]");
                }

                Console.WriteLine();
                Console.WriteLine("── 原始文件对照 ──");
                using (var o = DocXml.Load(env.OriginalDocx))
                {
                    Console.WriteLine("原始非空段(w:t+m:t): " + NonEmptyInclMath(o) +
                        " | w:t-only: " + o.NonEmptyParagraphCount() +
                        " | 表格: " + o.TableCount() + " | 图片: " + o.ImageCount());
                    Console.WriteLine("原始 w:t Unicode上下标字符=" + o.UnicodeMathCharsInPlainText());
                    int bold1e = 0;
                    foreach (var p in o.Paragraphs())
                    {
                        if (DocXml.PlainText(p).Trim().Length == 0) continue;
                        var rf = DocXml.FirstRunFonts(p);
                        if (rf != null && rf.Color == "1E1E1E" && rf.Bold == true) bold1e++;
                    }
                    Console.WriteLine("原始 1E1E1E 加粗首 run 段 = " + bold1e);
                }
            }
        }


        private static int NonEmptyInclMath(DocXml xml)
        {
            return xml.Paragraphs().Count(p =>
                (DocXml.PlainText(p) + DocXml.MathText(p)).Trim().Length > 0);
        }

        private static void Dump(DocXml xml, string prefix)
        {
            var p = xml.FindParagraphByPrefix(prefix);
            if (p == null) { Console.WriteLine("  [缺失] " + prefix); return; }
            DumpEl(xml, p, prefix.Substring(0, Math.Min(6, prefix.Length)));
        }

        private static void DumpEl(DocXml xml, System.Xml.Linq.XElement p, string label)
        {
            var rf = DocXml.FirstRunFonts(p);
            string text = DocXml.PlainText(p);
            bool inTbl = p.Ancestors(DocXml.W + "tbl").Any();
            Console.WriteLine("  [" + (inTbl ? "表" : "文") + "] " + text.Substring(0, Math.Min(16, text.Length)) +
                " → 东=" + (rf == null || rf.EastAsia == null ? "(继承)" : rf.EastAsia) +
                " 西=" + (rf == null || rf.Ascii == null ? "(继承)" : rf.Ascii) +
                " 号=" + (rf != null && rf.SizePt.HasValue ? rf.SizePt.Value.ToString() : "?") +
                " 粗=" + (rf != null && rf.Bold == true ? "是" : "否") +
                " 色=" + (rf != null ? rf.Color : "?") +
                "  数学区=" + p.Descendants(DocXml.M + "oMath").Count());
        }
    }
}
