using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace TypeRight.Acceptance
{
    /// <summary>
    /// OOXML 层探测：Word 文档在文件里是 w:r（run = 格式相同的连续文字）结构。
    /// 如果一段 177 字的正文只有十几个 run，那"按 run 分组"就比"逐字符 Range"快两个数量级。
    /// </summary>
    internal static class DiagRuns
    {
        public static int Run(string path, int maxParagraphs = 8)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine();
            Console.WriteLine("── OOXML run 结构探测 ──────────────────────");

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            using (var fs = File.OpenRead(path))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var entry = zip.GetEntry("word/document.xml");
                if (entry == null) { Console.WriteLine("[FAIL] 无 word/document.xml"); return 1; }

                XDocument xdoc;
                using (var s = entry.Open()) xdoc = XDocument.Load(s);

                int shown = 0;
                foreach (var p in xdoc.Descendants(w + "p"))
                {
                    var runs = new System.Collections.Generic.List<string>();
                    var runTexts = new System.Collections.Generic.List<string>();
                    foreach (var r in p.Descendants(w + "r"))
                    {
                        var t = r.Element(w + "t");
                        string txt = t != null ? t.Value : "";
                        if (txt.Length == 0) continue;
                        var rPr = r.Element(w + "rPr");
                        string ea = null, asc = null, sz = null;
                        if (rPr != null)
                        {
                            var f = rPr.Element(w + "rFonts");
                            if (f != null)
                            {
                                ea = (string)f.Attribute(w + "eastAsia");
                                asc = (string)f.Attribute(w + "ascii");
                            }
                            var szEl = rPr.Element(w + "sz");
                            if (szEl != null) sz = (string)szEl.Attribute(w + "val");
                        }
                        runs.Add(string.Format("Ea={0,-8} Ascii={1,-18} sz={2,-4} 文字={3}",
                            ea ?? "-", asc ?? "-", sz ?? "-",
                            txt.Length > 20 ? txt.Substring(0, 20) + "…" : txt));
                        runTexts.Add(txt);
                    }
                    if (runs.Count == 0) continue;

                    int charCount = 0;
                    foreach (string t in runTexts) charCount += t.Length;
                    if (charCount < 20) continue;   // 只看正文段

                    Console.WriteLine();
                    Console.WriteLine(string.Format("段落：{0} 字符 → {1} 个 run", charCount, runs.Count));
                    for (int i = 0; i < runs.Count && i < 12; i++)
                        Console.WriteLine("   run[{0}] {1}", i, runs[i]);
                    if (runs.Count > 12) Console.WriteLine("   … 还有 {0} 个 run", runs.Count - 12);

                    if (++shown >= maxParagraphs) break;
                }
            }
            return 0;
        }
    }
}
