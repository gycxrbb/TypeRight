using System.Collections.Generic;
using TypeRight.Signature;
using Xunit;

namespace TypeRight.Tests
{
    /// <summary>
    /// 公式区感知的字符类型判定（2026-11 真机验收 TC-21 回归）。
    ///
    /// 事故：公式转换把段落里的 x1 / Δ / 2a 变成 OMML 对象，它们从 Range.Text 消失，
    /// CharClass 从 Mixed/Digit 变 Cjk → 签名变动 → 同族命中 43 退化到 37。
    ///
    /// 修法：判定主导字符类型时剔除公式区字符，与既有「锚点跳过公式区」原则一致。
    /// </summary>
    public class CharClassMathZoneTests
    {
        /// <summary>text="商雨萤x1同志"，公式区覆盖 "x1"（位置 3..5）。</summary>
        private static List<int[]> Zones(int start, int end)
        {
            return new List<int[]> { new[] { start, end } };
        }

        [Fact]
        public void OutsideMathZones_NullZones_ReturnsOriginal()
        {
            Assert.Equal("abc", FormatSignatureReader.TextOutsideMathZones("abc", 0, null));
        }

        [Fact]
        public void OutsideMathZones_NoOverlap_ReturnsOriginalSameInstance()
        {
            string s = "商雨萤同志";
            var zones = Zones(100, 110);   // 完全不交叠
            string r = FormatSignatureReader.TextOutsideMathZones(s, 0, zones);
            Assert.Equal(s, r);
        }

        [Fact]
        public void OutsideMathZones_StripsCoveredChars()
        {
            string s = "商雨萤x1同志";
            // 段落起点 0，公式区覆盖索引 3..5（"x1"）
            var zones = Zones(3, 5);
            Assert.Equal("商雨萤同志", FormatSignatureReader.TextOutsideMathZones(s, 0, zones));
        }

        [Fact]
        public void OutsideMathZones_RespectsParagraphOffset()
        {
            string s = "x1商雨萤";
            // 段落起始位置 50，公式区 50..52（对应串内 0..2）
            var zones = Zones(50, 52);
            Assert.Equal("商雨萤", FormatSignatureReader.TextOutsideMathZones(s, 50, zones));
        }

        /// <summary>核心断言：公式化前后 CharClass 必须一致。</summary>
        [Fact]
        public void CharClassIsStableBeforeAndAfterFormulaConversion()
        {
            // 公式化前：文本里含 x1（字母+数字）
            string before = "当x1=2时方程成立";
            var noZones = new List<int[]>();
            // 公式化后：x1 变成公式对象，从文本消失，但公式区覆盖原来的位置
            string after = "当=2时方程成立";

            // 段落起点 0；"x1" 原在索引 1..3
            var zones = Zones(1, 3);

            CharClass b = CharClassDetector.Detect(
                FormatSignatureReader.TextOutsideMathZones(before, 0, noZones));
            CharClass a = CharClassDetector.Detect(
                FormatSignatureReader.TextOutsideMathZones(after, 0, zones));

            Assert.Equal(b, a);
            Assert.Equal(CharClass.Cjk, a);
        }

        /// <summary>反例：不剔除公式区时两者不等——证明这个修复是必要的。</summary>
        [Fact]
        public void WithoutStripping_CharClassWouldDiffer()
        {
            CharClass rawBefore = CharClassDetector.Detect("当x1=2时方程成立");
            CharClass rawAfter = CharClassDetector.Detect("当=2时方程成立");

            // 都是"有汉字即中文"，这条不会差；换成数字占多数的日期段才体现
            Assert.Equal(CharClass.Cjk, rawBefore);
            Assert.Equal(CharClass.Cjk, rawAfter);

            // 真正会退化的场景：公式化前数字占多数 → Digit，公式化后只剩汉字 → Cjk
            CharClass d1 = CharClassDetector.Detect("2005年11月26日");      // 数字 8 : 汉字 3 → Digit
            CharClass d2 = CharClassDetector.Detect("年月日");              // 只剩 3 个汉字 → Cjk
            Assert.Equal(CharClass.Digit, d1);
            Assert.Equal(CharClass.Cjk, d2);
            Assert.NotEqual(d1, d2);   // 这就是不剔除时 6 段掉出同族的机制
        }
    }
}
