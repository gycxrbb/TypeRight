using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>
    /// 字符类型过滤的真机验证：同一段正文分别"只动汉字"与"只动数字"，
    /// 检查目标类型被改、非目标类型保持原样。
    /// </summary>
    internal static class DiagClassFilter
    {
        private static int _fail;

        public static int Run(TestEnv env, string path)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine();
            Console.WriteLine("── 字符类型过滤验证 ──────────────────────");

            // 每次用例都用全新副本（要改格式）
            RunCase(env, path, CharClass.Cjk, "只动汉字 → 数字保持原样");
            RunCase(env, path, CharClass.Digit, "只动数字 → 汉字保持原样");
            RunCase(env, path, CharClass.None, "不过滤（旧行为）→ 整段都变");

            Console.WriteLine();
            Console.WriteLine(_fail == 0 ? "  字符类型过滤：ALL PASS" : "  字符类型过滤：FAIL " + _fail);
            return _fail;
        }

        private static void RunCase(TestEnv env, string path, CharClass target, string title)
        {
            Console.WriteLine();
            Console.WriteLine("[" + title + "]");

            string copy = System.IO.Path.Combine(env.OutputDir,
                "cls-" + target + "-" + Guid.NewGuid().ToString("N").Substring(0, 6) + ".docx");
            try { System.IO.File.Copy(path, copy, true); }
            catch (Exception ex) { Console.WriteLine("  [FAIL] 复制失败: " + ex.Message); _fail++; return; }

            Document doc = env.App.Documents.Open(copy, false, false, false);
            try
            {
                Paragraph para = null;
                string text = "";
                foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
                {
                    string t = (p.Range.Text ?? "").Replace("\r", "").Replace("\a", "");
                    int cjk = 0, dig = 0;
                    foreach (char c in t)
                    {
                        if (CharClassDetector.IsCjk(c)) cjk++;
                        else if (CharClassDetector.IsDigit(c)) dig++;
                    }
                    // 选同时含较多汉字和数字的正文段
                    if (cjk > 50 && dig > 5) { para = p; text = t; break; }
                }
                if (para == null) { Console.WriteLine("  [FAIL] 找不到合适的正文段"); _fail++; return; }

                int start = para.Range.Start;
                int firstCjk = -1, firstDigit = -1;
                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];
                    if (firstCjk < 0 && CharClassDetector.IsCjk(c)) firstCjk = i;
                    if (firstDigit < 0 && CharClassDetector.IsDigit(c)) firstDigit = i;
                    if (firstCjk >= 0 && firstDigit >= 0) break;
                }

                string cjkBefore = FontOf(doc, start + firstCjk);
                string digBefore = FontOf(doc, start + firstDigit);

                // 诊断：直接看纯逻辑层的切分结果
                var runs = CharClassDetector.SplitByClass(text);
                var keep = CharClassDetector.FilterRuns(runs, target == CharClass.None ? CharClass.Cjk : target, true);
                Console.WriteLine(string.Format("   [诊断] 段长={0} 切分={1} 片段 保留({2})={3} 片段",
                    text.Length, runs.Count, CharClassDetector.Name(target), keep.Count));
                if (keep.Count > 0)
                    Console.WriteLine(string.Format("   [诊断] 首个片段 Start={0} Len={1}", keep[0].Start, keep[0].Length));

                var units = new List<MatchUnit>
                {
                    new MatchUnit { Start = start, End = start + Math.Min(text.Length, para.Range.End - start - 1) }
                };
                var changes = new FormatChanges { SizePt = 20 };
                BatchFormatter.ApplyUnits(env.App, doc, units, changes, "test-" + target, true, target);

                string cjkAfter = FontOf(doc, start + firstCjk);
                string digAfter = FontOf(doc, start + firstDigit);

                Console.WriteLine("   汉字: " + cjkBefore + " → " + cjkAfter);
                Console.WriteLine("   数字: " + digBefore + " → " + digAfter);

                bool cjkChanged = cjkAfter != cjkBefore;
                bool digChanged = digAfter != digBefore;

                switch (target)
                {
                    case CharClass.Cjk:
                        Check(cjkChanged, "汉字被改动");
                        Check(!digChanged, "数字保持原样");
                        break;
                    case CharClass.Digit:
                        Check(digChanged, "数字被改动");
                        Check(!cjkChanged, "汉字保持原样");
                        break;
                    default:
                        Check(cjkChanged, "汉字被改动（不过滤）");
                        Check(digChanged, "数字也被改动（不过滤 = 旧行为）");
                        break;
                }
            }
            finally
            {
                try { doc.Close(false); } catch { }
                try { System.IO.File.Delete(copy); } catch { }
            }
        }

        private static string FontOf(Document doc, int pos)
        {
            try
            {
                Font f = doc.Range(pos, pos + 1).Font;
                return string.Format("{0}/{1}pt", f.NameFarEast, f.Size);
            }
            catch { return "(err)"; }
        }

        private static void Check(bool ok, string what)
        {
            if (ok) { Console.WriteLine("  [PASS] " + what); return; }
            Console.WriteLine("  [FAIL] " + what);
            _fail++;
        }
    }
}
