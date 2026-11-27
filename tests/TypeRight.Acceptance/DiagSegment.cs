using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>
    /// 验证「整段含数字就排除」这个方案在真实文档上的后果。
    /// 同时验证「按字符类型切分段落」的方案可行性与成本。
    /// </summary>
    internal static class DiagSegment
    {
        public static int Run(TestEnv env, string path)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Document doc = env.App.Documents.Open(path, false, true, false);
            try
            {
                Console.WriteLine();
                Console.WriteLine("── 方案对比：整段排除 vs 段内切分 ──────────────────────");

                // 收集所有"仿宋 14pt"正文段
                var body = new List<Paragraph>();
                var texts = new List<string>();
                var classes = new List<CharClass>();
                foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
                {
                    FormatSignature s = FormatSignatureReader.Read(p);
                    if (s.EastAsiaFont != "仿宋" || Math.Abs(s.SizePt - 14) > 0.01) continue;
                    string t = (p.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
                    if (t.Length == 0) continue;
                    body.Add(p); texts.Add(t); classes.Add(s.CharClass);
                }

                Console.WriteLine("正文段总数（仿宋 14pt）：" + body.Count);

                // ── 方案 A：整段含数字/字母就排除 ──
                int keptA = 0, droppedA = 0;
                Console.WriteLine();
                Console.WriteLine("[方案A] 整段含数字/字母 → 整段排除");
                for (int i = 0; i < body.Count; i++)
                {
                    bool hasAlnum = false;
                    foreach (char c in texts[i])
                        if (CharClassDetector.IsDigit(c) || CharClassDetector.IsLatin(c)) { hasAlnum = true; break; }
                    int cjk = 0, dig = 0;
                    foreach (char c in texts[i]) { if (CharClassDetector.IsCjk(c)) cjk++; else if (CharClassDetector.IsDigit(c)) dig++; }
                    if (hasAlnum) { droppedA++; Console.WriteLine(string.Format("   ✗ 排除 汉字{0,-4}数字{1,-3} {2}",
                        cjk, dig, texts[i].Length > 24 ? texts[i].Substring(0, 24) + "…" : texts[i])); }
                    else keptA++;
                }
                Console.WriteLine(string.Format("   → 保留 {0} 段，排除 {1} 段", keptA, droppedA));

                // ── 方案 B：段内按字符类型切分，只保留中文片段 ──
                Console.WriteLine();
                Console.WriteLine("[方案B] 段内按字符类型切分 → 只保留中文片段");
                var sw = System.Diagnostics.Stopwatch.StartNew();
                int totalSegs = 0, totalKeptChars = 0, rngCalls = 0;
                for (int i = 0; i < body.Count; i++)
                {
                    var segs = SplitByClass(texts[i]);
                    int pStart = body[i].Range.Start;
                    foreach (var seg in segs)
                    {
                        totalSegs++;
                        if (seg.Class != CharClass.Cjk) continue;   // 只保留中文片段
                        totalKeptChars += seg.Length;
                        // 实测：真的造 Range 要多久
                        Range r = doc.Range(pStart + seg.Start, pStart + seg.Start + seg.Length);
                        string dummy = r.Text;
                        rngCalls++;
                        if (dummy == null) Console.Write("");
                    }
                }
                sw.Stop();
                Console.WriteLine(string.Format("   → {0} 段、切出 {1} 个片段（造 {2} 次 Range）",
                    body.Count, totalSegs, rngCalls));
                Console.WriteLine(string.Format("   → 保留中文字符 {0} 个", totalKeptChars));
                Console.WriteLine(string.Format("   → 耗时 {0} ms（含逐片段造 Range + 读 Text）", sw.ElapsedMilliseconds));
                Console.WriteLine(string.Format("   → 平均每片段 {0:0.0} ms", (double)sw.ElapsedMilliseconds / Math.Max(1, rngCalls)));
            }
            finally { try { doc.Close(false); } catch { } }
            return 0;
        }

        private struct Seg { public int Start, Length; public CharClass Class; }

        /// <summary>把文本按字符类型切成连续片段（纯字符串运算，零 COM）。</summary>
        private static List<Seg> SplitByClass(string text)
        {
            var result = new List<Seg>();
            int i = 0;
            while (i < text.Length)
            {
                CharClass cls = ClassOf(text[i]);
                if (cls == CharClass.None) { i++; continue; }   // 标点/空白跳过
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
