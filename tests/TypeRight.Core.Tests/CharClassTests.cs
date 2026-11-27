using TypeRight.Signature;
using Xunit;

namespace TypeRight.Tests
{
    /// <summary>主导字符类型判定（同类选择新增维度：中文/字母/数字）。
    /// 口径：按各类字符"数量占比"取主导者，并列最多 → Mixed。</summary>
    public class CharClassTests
    {
        // ── 纯类型 ──

        [Theory]
        [InlineData("这是中文正文段落")]
        [InlineData("第")]
        [InlineData("一元二次方程")]
        public void Detect_PureChinese_IsCjk(string text)
        {
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect(text));
        }

        [Theory]
        [InlineData("abcdefgh")]
        [InlineData("Python")]
        [InlineData("ＡＢＣ")]
        [InlineData("ｘｙｚ")]
        public void Detect_PureLatin_IsLatin(string text)
        {
            Assert.Equal(CharClass.Latin, CharClassDetector.Detect(text));
        }

        [Theory]
        [InlineData("1234567890")]
        [InlineData("２０２５")]
        [InlineData("3.14159")]
        public void Detect_PureDigit_IsDigit(string text)
        {
            Assert.Equal(CharClass.Digit, CharClassDetector.Detect(text));
        }

        // ── 混排：按数量取主导 ──

        [Fact]
        public void Detect_ChineseWithFewDigits_IsStillCjk()
        {
            // 用户确认的口径：第1题 → 以中文为主
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect("第1题下列说法正确的是"));
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect("2025年预备党员"));
        }

        [Fact]
        public void Detect_ChineseWithFewLatin_IsStillCjk()
        {
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect("Python编程基础教程"));
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect("使用 Word 排版"));
        }

        [Fact]
        public void Detect_LatinWithFewDigits_IsLatin()
        {
            // 口径（2026-11 按真实文档校准，见 AGENTS.md）：无汉字时，拉丁与数字**同权**比较，
            // 平局即 Mixed。旧期望把 "ABC123" 判 Latin，等于让数字单方面从属于字母，
            // 那样纯数字段会被卷进字母组一起改格式——与实际要区分的目标相反。
            Assert.Equal(CharClass.Mixed, CharClassDetector.Detect("ABC123"));   // 3 字母 : 3 数字，平局
            Assert.Equal(CharClass.Latin, CharClassDetector.Detect("x1y2z3a"));   // 4 字母 : 3 数字，字母多数
            Assert.Equal(CharClass.Latin, CharClassDetector.Detect("abc"));      // 纯拉丁
        }

        [Fact]
        public void Detect_DigitWithFewLatin_IsDigit()
        {
            Assert.Equal(CharClass.Digit, CharClassDetector.Detect("123456a"));
            Assert.Equal(CharClass.Digit, CharClassDetector.Detect("2025-11-19"));
        }

        // ── 无主导者 → Mixed ──

        [Theory]
        [InlineData("AB12")]        // 2 字母 2 数字
        [InlineData("abc123")]
        [InlineData("中文")]        // 2 汉字 —— 仍应 Cjk（单一类型）
        public void Detect_TieOrSingle_FollowsRule(string text)
        {
            // AB12/abc123 并列 → Mixed；"中文"只有汉字 → Cjk
            CharClass expected = text == "中文" ? CharClass.Cjk : CharClass.Mixed;
            Assert.Equal(expected, CharClassDetector.Detect(text));
        }

        [Fact]
        public void Detect_ChineseAndLatinTie_IsMixed()
        {
            // 口径（2026-11 校准，见 AGENTS.md）：**有汉字即判中文**，不按字符数公平比较。
            // "Python编程基础教程" 6 汉字 vs 6 字母若判 Mixed，与用户"这是中文正文"的直觉相悖——
            // 中文夹英文词是常态，英文词不该让整段脱离中文组。
            // 只有"数字压过汉字"才例外（日期/数据段，见下一条用例）。
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect("中A"));
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect("你好ab"));
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect("Python编程基础教程"));
        }

        // ── 空/无有效字符 ──

        [Theory]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("   ")]
        [InlineData("\r\a")]
        [InlineData("，。！？；：")]
        [InlineData("——")]
        public void Detect_NoMeaningfulChars_IsNone(string text)
        {
            Assert.Equal(CharClass.None, CharClassDetector.Detect(text));
        }

        [Fact]
        public void Detect_PunctuationIgnored_OnlyContentCounts()
        {
            // 标点不参与计数：中文 + 大量标点仍判中文
            Assert.Equal(CharClass.Cjk, CharClassDetector.Detect("好，。！？、；："));
            // 纯标点 → None
            Assert.Equal(CharClass.None, CharClassDetector.Detect("【】《》（）"));
        }

        [Fact]
        public void Detect_FullWidthAndHalfWidthTreatedEqually()
        {
            Assert.Equal(CharClass.Digit, CharClassDetector.Detect("１２３"));
            Assert.Equal(CharClass.Digit, CharClassDetector.Detect("123"));
            Assert.Equal(CharClass.Latin, CharClassDetector.Detect("ＡＢＣ"));
            Assert.Equal(CharClass.Latin, CharClassDetector.Detect("ABC"));
        }

        // ── 字符判定单元 ──

        [Theory]
        [InlineData('中', true)]
        [InlineData('一', true)]
        [InlineData('\u3400', true)]   // 扩展 A
        [InlineData('\uF900', true)]   // 兼容区
        [InlineData('a', false)]
        [InlineData('1', false)]
        [InlineData('，', false)]
        public void IsCjk_DistinguishesCorrectly(char c, bool expected)
        {
            Assert.Equal(expected, CharClassDetector.IsCjk(c));
        }

        [Theory]
        [InlineData('a', true)]
        [InlineData('Z', true)]
        [InlineData('Ａ', true)]
        [InlineData('中', false)]
        [InlineData('1', false)]
        public void IsLatin_DistinguishesCorrectly(char c, bool expected)
        {
            Assert.Equal(expected, CharClassDetector.IsLatin(c));
        }

        [Theory]
        [InlineData('0', true)]
        [InlineData('９', true)]
        [InlineData('a', false)]
        [InlineData('中', false)]
        public void IsDigit_DistinguishesCorrectly(char c, bool expected)
        {
            Assert.Equal(expected, CharClassDetector.IsDigit(c));
        }

        // ── 真实场景 ──

        [Theory]
        [InlineData("这是第 1 行同类段落，用于取证。", CharClass.Cjk)]
        [InlineData("1234567890123456789", CharClass.Digit)]
        [InlineData("abcdefghijklmnopqrst", CharClass.Latin)]
        [InlineData("第1题下列说法正确的是", CharClass.Cjk)]
        [InlineData("Python编程基础教程内容", CharClass.Cjk)]
        [InlineData("2026 年 11 月 19 日", CharClass.Digit)]
        public void Detect_RealWorldSamples(string text, CharClass expected)
        {
            Assert.Equal(expected, CharClassDetector.Detect(text));
        }
    }
}
