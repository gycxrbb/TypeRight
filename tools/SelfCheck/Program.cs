using System;
using System.Collections.Generic;
using TypeRight.Signature;

namespace TypeRight.SelfCheck
{
    /// <summary>
    /// 独立自检程序（不依赖 xunit 测试宿主）。
    /// 用途：当 dotnet test 的 xunit 适配器在个别机器上加载失败时，
    /// 仍能对纯逻辑层做回归验证——CI/本机都能跑，退出码 = 失败数。
    /// 覆盖：CharClass 主导字符类型判定（2026-11 新增维度）。
    /// </summary>
    internal static class Program
    {
        private static int _fail;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("── CharClass 主导字符类型自检 ──────────────────");

            // 纯类型
            Expect("这是中文正文段落", CharClass.Cjk, "纯中文");
            Expect("第", CharClass.Cjk, "单个汉字");
            Expect("abcdefgh", CharClass.Latin, "纯拉丁");
            Expect("ＡＢＣ", CharClass.Latin, "全角拉丁");
            Expect("1234567890", CharClass.Digit, "纯数字");
            Expect("２０２５", CharClass.Digit, "全角数字");
            Expect("3.14159", CharClass.Digit, "含小数点数字");

            // 混排：按占比取主导（用户确认口径）
            Expect("第1题下列说法正确的是", CharClass.Cjk, "中文为主+1数字");
            Expect("2025年预备党员", CharClass.Cjk, "中文为主+数字年份");
            Expect("Python编程基础教程", CharClass.Cjk, "中文为主+英文词");
            Expect("使用 Word 排版", CharClass.Cjk, "中文为主+空格英文");
            Expect("ABC123", CharClass.Mixed, "字母数字等量 → 混合");
            Expect("123456a", CharClass.Digit, "数字明显占多数");

            // 并列 → Mixed（仅限无汉字时；有汉字时见下）
            Expect("AB12", CharClass.Mixed, "字母数字并列");
            Expect("你好ab", CharClass.Cjk, "汉字压过字母 → 中文");
            Expect("中A", CharClass.Cjk, "汉字压过字母 → 中文");

            // ── 真实文档校准（2026-11，考察报告-商雨萤.docx）──
            // 该文档正文段（汉字114/数字49）与日期段（汉字3/数字8）字体字号完全相同，
            // 唯一能区分的就是字符组成。数字压过汉字 → 数字段，避免日期被卷进正文组。
            Expect("2026年11月26日", CharClass.Digit, "日期段：数字压过汉字 → 数字");
            Expect("2005年5月出生", CharClass.Digit, "日期短语：数字7 汉字3 → 数字");
            Expect("商雨萤同志，女，2005年5月出生，2023级数据科学与大数据技术", CharClass.Cjk,
                "正文段：汉字远多于数字 → 中文");
            Expect("预备期内，党支部按照要求开展工作，共召开会议12次", CharClass.Cjk,
                "正文夹少量数字 → 中文");

            // 无有效字符 → None
            Expect("", CharClass.None, "空串");
            Expect(null, CharClass.None, "null");
            Expect("   ", CharClass.None, "纯空白");
            Expect("，。！？；：", CharClass.None, "纯标点");
            Expect("【】《》（）", CharClass.None, "纯括号");

        // 真实场景（与验收探针一致）
        Expect("这是第 1 行同类段落，用于取证。", CharClass.Cjk, "真实中文段");
            Expect("1234567890123456789", CharClass.Digit, "真实数字段");
            Expect("abcdefghijklmnopqrst", CharClass.Latin, "真实字母段");

            SplitChecks();

            Console.WriteLine();
            Console.WriteLine(_fail == 0
                ? "ALL PASS（CharClass 自检全部通过）"
                : "FAIL " + _fail + " 项");
            return _fail;
        }

        // ── 切分/过滤（同类选择"汉字归汉字、数字归数字"的底层）──

        private static void SplitChecks()
        {
            Console.WriteLine();
            Console.WriteLine("── 按字符类型切分/过滤 ──────────────────");

            var runs = CharClassDetector.SplitByClass("商雨萤同志，女，2005年");
            ExpectInt(runs.Count, 4, "切分段数：中文/中文/数字/中文");

            var keepCjk = CharClassDetector.FilterRuns(runs, CharClass.Cjk, true);
            ExpectInt(keepCjk.Count, 2, "合并后中文片段数");
            if (keepCjk.Count > 0)
                ExpectStr(keepCjk[0].Start + "," + keepCjk[0].Length, "0,7", "首个中文片段 商雨萤同志，女");

            var keepDig = CharClassDetector.FilterRuns(runs, CharClass.Digit, true);
            ExpectInt(keepDig.Count, 1, "数字片段数");
            if (keepDig.Count > 0)
                ExpectStr("2005", TextOf("商雨萤同志，女，2005年", keepDig[0]), "数字片段内容");

            // 核心回归：多个中文片段被数字隔开，必须全部保留（不能只剩最后一个）
            var multi = CharClassDetector.SplitByClass("甲乙1丙丁2戊");
            var multiKeep = CharClassDetector.FilterRuns(multi, CharClass.Cjk, false);
            ExpectInt(multiKeep.Count, 3, "三个中文片段全部保留（回归：曾丢失）");

            // 真实段落：中文片段必须远多于 1
            string real = "商雨萤同志，女，2005年5月出生，2023级数据科学与大数据技术专业本科生。";
            var realRuns = CharClassDetector.SplitByClass(real);
            var realKeep = CharClassDetector.FilterRuns(realRuns, CharClass.Cjk, true);
            ExpectBool(realKeep.Count >= 4, "真实段落中文片段 >= 4（实际 " + realKeep.Count + "）");

            // 边界：空、纯标点
            ExpectInt(CharClassDetector.SplitByClass("").Count, 0, "空串零片段");
            ExpectInt(CharClassDetector.SplitByClass("，。！？").Count, 0, "纯标点零片段");
        }

        private static string TextOf(string src, CharClassDetector.Run r)
        {
            if (r.Start < 0 || r.End > src.Length) return "(越界)";
            return src.Substring(r.Start, r.Length);
        }

        private static void ExpectInt(int actual, int expected, string what)
        {
            if (actual == expected) Console.WriteLine("  [PASS] " + what + " = " + actual);
            else { Console.WriteLine("  [FAIL] " + what + " | 期望 " + expected + " 实际 " + actual); _fail++; }
        }

        private static void ExpectStr(string actual, string expected, string what)
        {
            if (actual == expected) Console.WriteLine("  [PASS] " + what + " = " + actual);
            else { Console.WriteLine("  [FAIL] " + what + " | 期望 " + expected + " 实际 " + actual); _fail++; }
        }

        private static void ExpectBool(bool ok, string what)
        {
            if (ok) Console.WriteLine("  [PASS] " + what);
            else { Console.WriteLine("  [FAIL] " + what); _fail++; }
        }

        private static void Expect(string text, CharClass expected, string what)
        {
            CharClass actual = CharClassDetector.Detect(text);
            if (actual == expected)
            {
                Console.WriteLine("  [PASS] " + what + " → " + CharClassDetector.Name(actual));
            }
            else
            {
                Console.WriteLine("  [FAIL] " + what + " | 期望 " + expected + " 实际 " + actual);
                _fail++;
            }
        }
    }
}
