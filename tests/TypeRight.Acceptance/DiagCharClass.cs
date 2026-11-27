using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>
    /// 字符类型维度探针（2026-11）：
    /// 验证"中文正文 / 纯数字 / 纯字母 / 中英混排"在**当前签名判定**下是否被判为同类。
    /// 结论直接决定是否需要新增"字符类型"维度。
    /// </summary>
    internal static class DiagCharClass
    {
        public static int Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine();
            Console.WriteLine("── 字符类型维度探针 ──────────────────────");

            Document doc = env.App.Documents.Add();
            try
            {
                // 全部统一：宋体 + Times New Roman + 10.5pt，唯一变量是"内容由什么字符组成"
                string[] texts =
                {
                    "这是中文正文段落内容测试题",
                    "1234567890123456789",
                    "abcdefghijklmnopqrst",
                    "第1题下列说法正确的是",
                    "Python编程基础教程内容",
                    "第2题下列说法正确的是"
                };
                string[] labels = { "纯中文A", "纯数字", "纯字母", "中+数字", "字母+中", "纯中文B" };

                foreach (string t in texts) AddPara(doc, t);

                Console.WriteLine();
                Console.WriteLine("序号 类型        首字符  尾字符  Ea字体   Latin字体        字号   加粗");
                Console.WriteLine("──────────────────────────────────────────────────────────────");
                var sigs = new List<FormatSignature>();
                for (int i = 1; i <= texts.Length; i++)
                {
                    Paragraph p = doc.Paragraphs[i];
                    FormatSignature s = FormatSignatureReader.Read(p);
                    sigs.Add(s);
                    string raw = (p.Range.Text ?? "").Replace("\r", "").Replace("\a", "");
                    string first = raw.Length > 0 ? raw[0].ToString() : "?";
                    string last = raw.Length > 0 ? raw[raw.Length - 1].ToString() : "?";
                    Console.WriteLine(string.Format("{0,-4} {1,-10} {2,-6} {3,-6} {4,-8} {5,-16} {6,-6} B{7}",
                        i, labels[i - 1], first, last,
                        s.EastAsiaFont ?? "?", s.LatinFont ?? "?",
                        double.IsNaN(s.SizePt) ? "?" : s.SizePt.ToString("0.#"),
                        s.BoldRaw == -1 ? 1 : 0));
                }

                // 两两对比：谁和谁被判同类（严格模式）
                Console.WriteLine();
                Console.WriteLine("严格模式两两判定：");
                var sameClass = new List<string>();
                for (int i = 0; i < sigs.Count; i++)
                {
                    for (int j = i + 1; j < sigs.Count; j++)
                    {
                        if (sigs[i].Matches(sigs[j], FormatSignature.IgnoreDims.None))
                        {
                            sameClass.Add(labels[i] + " ≡ " + labels[j]);
                            Console.WriteLine("  ★ 同类: " + labels[i] + "  ≡  " + labels[j]);
                        }
                    }
                }

                // ── 断言（2026-11 新增字符类型维度）──
                Console.WriteLine();
                bool pureDigitIsolated = sigs[1].CharClass == CharClass.Digit;
                bool pureLatinIsolated = sigs[2].CharClass == CharClass.Latin;
                Check(pureDigitIsolated, "纯数字段判定为 Digit", "Digit",
                    CharClassDetector.Name(sigs[1].CharClass));
                Check(pureLatinIsolated, "纯字母段判定为 Latin", "Latin",
                    CharClassDetector.Name(sigs[2].CharClass));

                // 纯数字/纯字母不得与任何中文段互判同类
                bool digitLeaks = false, latinLeaks = false;
                for (int j = 0; j < sigs.Count; j++)
                {
                    if (j == 1) continue;
                    if (sigs[1].Matches(sigs[j], FormatSignature.IgnoreDims.None)) digitLeaks = true;
                }
                for (int j = 0; j < sigs.Count; j++)
                {
                    if (j == 2) continue;
                    if (sigs[2].Matches(sigs[j], FormatSignature.IgnoreDims.None)) latinLeaks = true;
                }
                Check(!digitLeaks, "纯数字段不与任何其他段同类（原 5 组 → 应 0）", "0 组", "有泄漏");
                Check(!latinLeaks, "纯字母段不与任何其他段同类（原 5 组 → 应 0）", "0 组", "有泄漏");

                // 中文段之间仍要能互判（不能因为加维度把正常中文匹配弄丢）
                bool cjkStillMatch = sigs[3].Matches(sigs[5], FormatSignature.IgnoreDims.None);
                Check(cjkStillMatch, "中+数字 与 纯中文 仍判同类（中文为主）", "同类", "不同类");

                // 宽松模式忽略该维度后，纯数字与纯字母应可互判（维度确实可关）
                bool ignorable = sigs[1].Matches(sigs[2],
                    FormatSignature.IgnoreDims.CharClass | FormatSignature.IgnoreDims.SizePt);
                Check(ignorable, "勾选忽略字符类型后 纯数字 ≡ 纯字母", "同类", "不同类");

                Console.WriteLine();
                Console.WriteLine(fail == 0 ? "  字符类型维度：ALL PASS" : "  字符类型维度：FAIL " + fail);
                return fail;
            }
            finally { try { doc.Close(false); } catch { } }
        }

        private static int fail;

        private static void Check(bool ok, string what, string expected, string actual)
        {
            if (ok) { Console.WriteLine("  [PASS] " + what); return; }
            Console.WriteLine("  [FAIL] " + what + " | 期望 " + expected + " 实际 " + actual);
            fail++;
        }

        private static void AddPara(Document doc, string text)
        {
            doc.Content.InsertAfter(text + "\r");
            Paragraph p = doc.Paragraphs[doc.Paragraphs.Count];
            Range r = p.Range;
            r.Font.NameFarEast = "宋体";
            r.Font.NameAscii = "Times New Roman";
            r.Font.NameOther = "Times New Roman";
            r.Font.Size = 10.5f;
            try { r.Font.Bold = 0; } catch { }
        }
    }
}
