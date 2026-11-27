using System;
using TypeRight.Formatting;
using Xunit;

namespace TypeRight.Tests
{
    /// <summary>字号手输解析（浮条字号框可打字录入：四号 / 小四 / 14 / 12.5 / 14磅）。</summary>
    public class FontSizeParseTests
    {
        [Theory]
        [InlineData("四号", 14)]
        [InlineData("小四", 12)]
        [InlineData("五号", 10.5)]
        [InlineData("初号", 42)]
        [InlineData("八号", 5)]
        [InlineData("三号", 16)]
        public void TryParse_AcceptsChineseSizeNames(string input, double expected)
        {
            Assert.Equal(expected, ChineseFontSizes.TryParse(input).Value, 3);
        }

        [Theory]
        [InlineData("14", 14)]
        [InlineData("12.5", 12.5)]
        [InlineData("10.5", 10.5)]
        [InlineData("72", 72)]
        [InlineData("20", 20)]
        [InlineData("1", 1)]
        public void TryParse_AcceptsNumericPointSizes(string input, double expected)
        {
            Assert.Equal(expected, ChineseFontSizes.TryParse(input).Value, 3);
        }

        [Theory]
        [InlineData("14磅")]
        [InlineData("14 磅")]
        [InlineData("14pt")]
        [InlineData("14 PT")]
        [InlineData("12.5磅")]
        public void TryParse_StripsPtSuffixes(string input)
        {
            Assert.NotNull(ChineseFontSizes.TryParse(input));
            Assert.True(ChineseFontSizes.TryParse(input).Value > 10);
        }

        [Fact]
        public void TryParse_HandlesFullWidthDigitsFromChineseIme()
        {
            // 中文输入法下敲的是全角数字：１４、１０．５
            Assert.Equal(14, ChineseFontSizes.TryParse("１４").Value, 3);
            Assert.Equal(10.5, ChineseFontSizes.TryParse("１０．５").Value, 3);
        }

        [Fact]
        public void TryParse_TrimsSurroundingWhitespace()
        {
            Assert.Equal(14, ChineseFontSizes.TryParse("  14  ").Value, 3);
            Assert.Equal(14, ChineseFontSizes.TryParse("四号 ").Value, 3);
            Assert.Equal(14, ChineseFontSizes.TryParse("\u3000四号").Value, 3);   // 全角空格
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        [InlineData("abc")]
        [InlineData("十四号")]
        [InlineData("号")]
        [InlineData("12.5.6")]
        [InlineData("-14")]
        public void TryParse_RejectsGarbage(string input)
        {
            Assert.Null(ChineseFontSizes.TryParse(input));
        }

        [Theory]
        [InlineData("0")]
        [InlineData("0.5")]
        [InlineData("1639")]
        [InlineData("9999")]
        public void TryParse_RejectsOutOfWordRange(string input)
        {
            // Word 字号范围 1~1638；越界直接判非法，不悄悄截断
            Assert.Null(ChineseFontSizes.TryParse(input));
        }

        [Theory]
        [InlineData(1)]
        [InlineData(1638)]
        public void TryParse_AcceptsRangeBoundaries(double pt)
        {
            Assert.Equal(pt, ChineseFontSizes.TryParse(pt.ToString("0.#")).Value, 3);
        }

        [Fact]
        public void TryParse_RoundTripsEveryChineseName()
        {
            foreach (var e in ChineseFontSizes.Table)
            {
                Assert.Equal(e.Pt, ChineseFontSizes.TryParse(e.Name).Value, 3);
            }
        }

        [Fact]
        public void TryParse_ResultFeedsLabelConsistently()
        {
            // 打了「四号」→ 解析成 14 → 显示回「四号」，来回一致
            double? pt = ChineseFontSizes.TryParse("四号");
            Assert.NotNull(pt);
            Assert.Equal("四号", ChineseFontSizes.Label(pt.Value));
        }

        [Fact]
        public void TryParse_ToleratesTrailingCharsAfterChineseName()
        {
            // 手输残留（「四号磅」）不应整条输入报废
            Assert.Equal(14, ChineseFontSizes.TryParse("四号磅").Value, 3);
        }
    }
}
