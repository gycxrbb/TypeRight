using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Acceptance
{
    /// <summary>验收环境：资产路径、Word 实例、断言收集（§8：fail 必须带期望/实际/锚点）。</summary>
    public sealed class TestEnv
    {
        public Application App;
        public readonly List<string> Failures = new List<string>();
        private string _currentCase = "";

        public readonly string RepoRoot;
        public readonly string TestsDir;
        public readonly string OriginalDocx;
        public readonly string GoldenDocx;
        public readonly string OutputDir;

        public TestEnv(Application app)
        {
            App = app;
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            while (dir != null && !File.Exists(Path.Combine(dir, "TypeRight.sln")))
                dir = Path.GetDirectoryName(dir);
            RepoRoot = dir ?? AppDomain.CurrentDomain.BaseDirectory;
            TestsDir = Path.Combine(RepoRoot, "tests");
            OriginalDocx = Path.Combine(TestsDir, "初三数学代数终极讲义_一元二次方程与二次函数大一统.docx");
            GoldenDocx = Path.Combine(TestsDir, "初三数学代数终极讲义_字体统一版.docx");
            OutputDir = Path.Combine(TestsDir, "output");
            Directory.CreateDirectory(OutputDir);
        }

        public void Begin(string testCase)
        {
            _currentCase = testCase;
            Console.WriteLine();
            Console.WriteLine("── " + testCase + " ──────────────────────");
        }

        public void Check(bool condition, string what, string expected = null, string actual = null, string anchor = null)
        {
            if (condition)
            {
                Console.WriteLine("  [PASS] " + what);
            }
            else
            {
                Console.WriteLine("  [FAIL] " + what);
                if (expected != null || actual != null)
                {
                    Console.WriteLine("         期望: " + (expected ?? "(null)"));
                    Console.WriteLine("         实际: " + (actual ?? "(null)"));
                }
                if (anchor != null) Console.WriteLine("         锚点: " + anchor);
                Failures.Add(_currentCase + " | " + what + " | 期望 " + (expected ?? "?") + " 实际 " + (actual ?? "?") + (anchor != null ? " @ " + anchor : ""));
            }
        }

        /// <summary>从原始资产复制全新副本并打开（禁止改动原件）。</summary>
        public Document OpenCopy(string tag)
        {
            string target = Path.Combine(OutputDir, tag + ".docx");
            File.Copy(OriginalDocx, target, true);
            return App.Documents.Open(target);
        }

        public string CopyPath(string tag)
        {
            return Path.Combine(OutputDir, tag + ".docx");
        }

        /// <summary>回收 Word 进程（重型用例间调用，防止长跑 RPC 崩溃/内存膨胀）。</summary>
        public void RecycleWord()
        {
            try { App.Quit(); App.Dispose(); } catch { }
            System.Threading.Thread.Sleep(1500);
            App = new Application
            {
                Visible = false,
                DisplayAlerts = NetOffice.WordApi.Enums.WdAlertLevel.wdAlertsNone
            };
        }

        public Paragraph FindParagraph(Document doc, string prefix)
        {
            foreach (Paragraph p in doc.Paragraphs)
            {
                try
                {
                    string t = (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v');
                    if (t.StartsWith(prefix, StringComparison.Ordinal)) return p;
                }
                catch { }
            }
            return null;
        }

        public List<Paragraph> ParagraphsStartingWith(Document doc, string prefix)
        {
            var list = new List<Paragraph>();
            foreach (Paragraph p in doc.Paragraphs)
            {
                try
                {
                    string t = (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v');
                    if (t.StartsWith(prefix, StringComparison.Ordinal)) list.Add(p);
                }
                catch { }
            }
            return list;
        }
    }

    /// <summary>docx 文件级断言（Open XML SDK，规格 §6 指引）。</summary>
    public sealed class DocXml : IDisposable
    {
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

        private readonly WordprocessingDocument _doc;
        public readonly XElement Root;

        private DocXml(WordprocessingDocument doc)
        {
            _doc = doc;
            Root = System.Xml.Linq.XElement.Parse(doc.MainDocumentPart.Document.OuterXml);
        }

        public static DocXml Load(string path)
        {
            return new DocXml(WordprocessingDocument.Open(path, false));
        }

        public void Dispose() { _doc.Dispose(); }

        public static string PlainText(XElement el)
        {
            var sb = new StringBuilder();
            foreach (var t in el.Descendants(W + "t")) sb.Append(t.Value);
            return sb.ToString();
        }

        public static string MathText(XElement el)
        {
            var sb = new StringBuilder();
            foreach (var t in el.Descendants(M + "t")) sb.Append(t.Value);
            return sb.ToString();
        }

        public List<XElement> Paragraphs()
        {
            return Root.Descendants(W + "p").ToList();
        }

        public int NonEmptyParagraphCount()
        {
            // 公式段转换后文本进 m:t（w:t 为空），计数需含公式文本（黄金取证：145 段）
            return Paragraphs().Count(p => (PlainText(p) + MathText(p)).Trim().Length > 0);
        }

        public int TableCount() { return Root.Descendants(W + "tbl").Count(); }

        public int ImageCount() { return _doc.MainDocumentPart.ImageParts.Count(); }

        public int OMathCount() { return Root.Descendants(M + "oMath").Count(); }

        public int MathElementCount(string localName) { return Root.Descendants(M + localName).Count(); }

        private const string SupSubChars = "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎√";

        public int UnicodeMathCharsInPlainText()
        {
            int count = 0;
            foreach (var t in Root.Descendants(W + "t"))
                foreach (char c in t.Value)
                    if (SupSubChars.IndexOf(c) >= 0) count++;
            return count;
        }

        public XElement FindParagraphByPrefix(string prefix)
        {
            foreach (XElement p in Paragraphs())
            {
                if (PlainText(p).TrimStart().StartsWith(prefix, StringComparison.Ordinal)) return p;
            }
            return null;
        }

        public List<XElement> FindParagraphsByPrefix(string prefix)
        {
            return Paragraphs()
                .Where(p => PlainText(p).TrimStart().StartsWith(prefix, StringComparison.Ordinal))
                .ToList();
        }

        public XElement FindTableCellContaining(string contains)
        {
            string compact = contains.Replace(" ", "");
            foreach (XElement tc in Root.Descendants(W + "tc"))
            {
                if (PlainText(tc).Replace(" ", "").Contains(compact)) return tc;
            }
            return null;
        }

        /// <summary>首个含可见文本的 run 的字体信息：eastAsia/ascii/字号pt/加粗/颜色。</summary>
        public sealed class RunFonts
        {
            public string EastAsia;
            public string Ascii;
            public double? SizePt;
            public bool? Bold;
            public string Color;
        }

        public static RunFonts FirstRunFonts(XElement paragraph)
        {
            foreach (XElement r in paragraph.Descendants(W + "r"))
            {
                string text = PlainText(r);
                if (text.Trim().Length == 0) continue;

                var rf = new RunFonts();
                XElement rPr = r.Element(W + "rPr");
                if (rPr != null)
                {
                    XElement fonts = rPr.Element(W + "rFonts");
                    if (fonts != null)
                    {
                        rf.EastAsia = (string)fonts.Attribute(W + "eastAsia");
                        rf.Ascii = (string)fonts.Attribute(W + "ascii");
                    }
                    XElement sz = rPr.Element(W + "sz");
                    if (sz != null)
                    {
                        double half;
                        if (double.TryParse((string)sz.Attribute(W + "val"), out half)) rf.SizePt = half / 2.0;
                    }
                    XElement b = rPr.Element(W + "b");
                    if (b != null)
                    {
                        string v = (string)b.Attribute(W + "val");
                        rf.Bold = v == "0" || v == "false" ? false : true;
                    }
                    XElement c = rPr.Element(W + "color");
                    if (c != null) rf.Color = (string)c.Attribute(W + "val");
                }
                return rf;
            }
            return null;
        }

        /// <summary>按首 run 中文字体统计段数（楷体 8 / 黑体 16 / 宋体 71 计数断言）。</summary>
        public int CountParagraphsByEastAsia(string eastAsia)
        {
            int n = 0;
            foreach (XElement p in Paragraphs())
            {
                if (PlainText(p).Trim().Length == 0) continue;
                RunFonts rf = FirstRunFonts(p);
                if (rf != null && string.Equals(rf.EastAsia, eastAsia, StringComparison.Ordinal)) n++;
            }
            return n;
        }

        /// <summary>全文 CJK 字符流（顺序保留），TC-13 内容零损失对比。</summary>
        public string CjkStream()
        {
            var sb = new StringBuilder();
            foreach (var t in Root.Descendants(W + "t"))
                foreach (char c in t.Value)
                    if (c >= '㐀' && c <= '鿿') sb.Append(c);
            return sb.ToString();
        }
    }

    /// <summary>OMML 结构锚点辅助（规格 §4.3）。</summary>
    public static class OmmlAssert
    {
        public static List<XElement> Fractions(XElement scope)
        {
            return scope.Descendants(DocXml.M + "f").ToList();
        }

        public static XElement Num(XElement fraction) { return fraction.Element(DocXml.M + "num"); }
        public static XElement Den(XElement fraction) { return fraction.Element(DocXml.M + "den"); }

        public static bool DenTextIs(XElement fraction, string text)
        {
            XElement den = Den(fraction);
            return den != null && DocXml.MathText(den).Trim() == text;
        }

        public static bool ContainsRadicalWith(XElement scope, string radicandContains)
        {
            return scope.Descendants(DocXml.M + "rad")
                .Any(rad => DocXml.MathText(rad).Contains(radicandContains));
        }
    }
}
