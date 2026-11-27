using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>
    /// 澄清「字符类型」到底是不是"属性"。
    /// 核心演示：同一个字符的「文本内容」和「字体属性」是两个独立维度。
    /// </summary>
    internal static class DiagConcept
    {
        public static int Run(TestEnv env, string path)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.App.Documents.Open(path, false, true, false);
            try
            {
                Paragraph target = null;
                foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
                {
                    string t = (p.Range.Text ?? "").Replace("\r", "").Replace("\a", "");
                    if (t.Length > 150) { target = p; break; }
                }
                if (target == null) return 1;

                int start = target.Range.Start;
                string text = (target.Range.Text ?? "").Replace("\r", "").Replace("\a", "");

                Console.WriteLine();
                Console.WriteLine("══ 同一个字符，两个完全独立的维度 ══");
                Console.WriteLine();
                Console.WriteLine("维度1「文本内容」= 字符码本身（不碰 Word，纯字符串就能判断）");
                Console.WriteLine("维度2「字体属性」= Word 给它记的格式（要 COM 往返才能读）");
                Console.WriteLine();
                Console.WriteLine("字符   文本内容判断        字体属性(Word记的)");
                Console.WriteLine("──────────────────────────────────────────────────────────");

                int shown = 0;
                for (int i = 0; i < text.Length && shown < 14; i++)
                {
                    char c = text[i];
                    if (c == '\r' || c == '\a') continue;

                    // 维度1：纯字符串运算，零成本
                    string kind;
                    if (CharClassDetector.IsCjk(c)) kind = "汉字";
                    else if (CharClassDetector.IsDigit(c)) kind = "数字";
                    else if (CharClassDetector.IsLatin(c)) kind = "字母";
                    else kind = "标点/其他";

                    // 维度2：COM 往返
                    Range one = doc.Range(start + i, start + i + 1);
                    Font f = one.Font;
                    string font = string.Format("{0} / {1} / {2}pt", f.NameFarEast, f.NameAscii, f.Size);

                    Console.WriteLine(string.Format(" '{0}'   {1,-10}       {2}", c, kind, font));
                    shown++;
                }

                Console.WriteLine();
                Console.WriteLine("结论：");
                Console.WriteLine("  · 字符类型（汉字/数字/字母）是【文本自带的】，Word 不需要额外存");
                Console.WriteLine("  · 字体属性是【另存一份的格式】，同一个段里可以完全一样");
                Console.WriteLine("  · 所以「字体一样」≠「字符类型一样」——两者本来就是不同的东西");

                // ── 演示：段内切分后的实际片段 ──
                Console.WriteLine();
                Console.WriteLine("══ 若按字符类型切分，这段会变成哪些片段 ══");
                var segs = SplitByClass(text);
                int cjkChars = 0, digitChars = 0, latinChars = 0;
                foreach (var s in segs)
                {
                    if (s.Class == CharClass.Cjk) cjkChars += s.Length;
                    else if (s.Class == CharClass.Digit) digitChars += s.Length;
                    else if (s.Class == CharClass.Latin) latinChars += s.Length;
                }
                int shownSeg = 0;
                foreach (var s in segs)
                {
                    if (shownSeg++ >= 14) { Console.WriteLine("   … 共 " + segs.Count + " 个片段"); break; }
                    string sub = text.Substring(s.Start, Math.Min(s.Length, 18));
                    Console.WriteLine(string.Format("   [{0,-4}] {1,-14} {2}{3}",
                        CharClassDetector.Name(s.Class), s.Length + " 字",
                        sub, s.Length > 18 ? "…" : ""));
                }
                Console.WriteLine();
                Console.WriteLine(string.Format("   汉字 {0} 个 / 数字 {1} 个 / 字母 {2} 个 → 共 {3} 个片段",
                    cjkChars, digitChars, latinChars, segs.Count));
                Console.WriteLine();
                Console.WriteLine("   方案A（整段含数字就排除）：整段消失 → 这 " + cjkChars + " 个汉字一起没了 ✗");
                Console.WriteLine("   方案B（段内切分）：只选 " + cjkChars + " 个汉字，" + digitChars + " 个数字不动 ✓");
            }
            finally { try { doc.Close(false); } catch { } }
            return 0;
        }

        private struct Seg { public int Start, Length; public CharClass Class; }

        private static List<Seg> SplitByClass(string text)
        {
            var result = new List<Seg>();
            int i = 0;
            while (i < text.Length)
            {
                CharClass cls = ClassOf(text[i]);
                if (cls == CharClass.None) { i++; continue; }
                int j = i + 1;
                while (j < text.Length && ClassOf(text[j]) == cls) j++;
                result.Add(new Seg { Start = i, Length = j - i, Class = cls });
                i = j;
            }
            return result;
        }

        private static CharClass ClassOf(char c)
        {
            if (CharClassDetector.IsCjk(c)) return CharClass.Cjk;
            if (CharClassDetector.IsDigit(c)) return CharClass.Digit;
            if (CharClassDetector.IsLatin(c)) return CharClass.Latin;
            return CharClass.None;
        }
    }
}
