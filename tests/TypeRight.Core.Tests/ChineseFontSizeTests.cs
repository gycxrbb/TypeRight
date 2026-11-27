using System;
using TypeRight.Formatting;
using Xunit;

namespace TypeRight.Tests
{
    /// <summary>中文字号换算（浮条字号显示口径：四号 / 小四 / 五号…）。</summary>
    public class ChineseFontSizeTests
    {
        [Theory]
        [InlineData(42, "初号")]
        [InlineData(36, "小初")]
        [InlineData(26, "一号")]
        [InlineData(24, "小一")]
        [InlineData(22, "二号")]
        [InlineData(18, "小二")]
        [InlineData(16, "三号")]
        [InlineData(15, "小三")]
        [InlineData(14, "四号")]
        [InlineData(12, "小四")]
        [InlineData(10.5, "五号")]
        [InlineData(9, "小五")]
        [InlineData(7.5, "六号")]
        [InlineData(6.5, "小六")]
        [InlineData(5.5, "七号")]
        [InlineData(5, "八号")]
        public void NameOf_MapsEveryStandardChineseSize(double pt, string expected)
        {
            Assert.Equal(expected, ChineseFontSizes.NameOf(pt));
        }

        [Theory]
        [InlineData(13)]
        [InlineData(11)]
        [InlineData(12.5)]
        [InlineData(72)]
        [InlineData(1)]
        public void NameOf_ReturnsNullForNonStandardPointSizes(double pt)
        {
            Assert.Null(ChineseFontSizes.NameOf(pt));
        }

        [Fact]
        public void NameOf_ToleratesFloatingPointJitter()
        {
            // Word 读回的 10.5 可能带末位抖动；显示不应退化成分数
            Assert.Equal("五号", ChineseFontSizes.NameOf(10.5000001));
            Assert.Equal("五号", ChineseFontSizes.NameOf(10.4999999));
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        public void NameOf_ReturnsNullForUnknownSize(double pt)
        {
            Assert.Null(ChineseFontSizes.NameOf(pt));
        }

        [Theory]
        [InlineData("四号", 14)]
        [InlineData("小四", 12)]
        [InlineData("五号", 10.5)]
        [InlineData("初号", 42)]
        public void PtOf_MapsNameBackToPoints(string name, double expected)
        {
            Assert.Equal(expected, ChineseFontSizes.PtOf(name).Value, 3);
        }

        [Theory]
        [InlineData("四號")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("14")]
        public void PtOf_ReturnsNullForUnknownName(string name)
        {
            Assert.Null(ChineseFontSizes.PtOf(name));
        }

        [Fact]
        public void Constants_RoundTripEveryTableEntry()
        {
            foreach (var e in ChineseFontSizes.Table)
            {
                Assert.Equal(e.Name, ChineseFontSizes.NameOf(e.Pt));
                Assert.Equal(e.Pt, ChineseFontSizes.PtOf(e.Name).Value, 3);
            }
        }

        [Theory]
        [InlineData(14, "四号")]
        [InlineData(10.5, "五号")]
        [InlineData(12.5, "12.5")]
        [InlineData(13, "13")]
        public void Label_PrefersChineseNameAndFallsBackToNumber(double pt, string expected)
        {
            Assert.Equal(expected, ChineseFontSizes.Label(pt));
        }

        [Fact]
        public void Label_ShowsDashForUnknownSize()
        {
            Assert.Equal("—", ChineseFontSizes.Label(double.NaN));
        }

        [Theory]
        [InlineData(14, "四号（14 磅）")]
        [InlineData(12.5, "12.5 磅")]
        public void Description_ShowsNameAndPointsForTooltip(double pt, string expected)
        {
            Assert.Equal(expected, ChineseFontSizes.Description(pt));
        }

        [Fact]
        public void PtText_DropsTrailingZeroButKeepsHalfPoint()
        {
            Assert.Equal("14", ChineseFontSizes.PtText(14));
            Assert.Equal("10.5", ChineseFontSizes.PtText(10.5));
        }

        [Fact]
        public void Label_NeverRendersZeroForUnreadableSize()
        {
            // 回归：文本同类模式的 SampleSignature 是 default(struct)，SizePt 落在 0——
            // 浮条曾因此显示「0」。0 只会来自 default 结构体（Word 字号恒为正），
            // 必须按"未知"呈现，绝不能把 0 当成一个真实字号印在界面上。
            Assert.Equal("0", ChineseFontSizes.Label(0));          // 纯换算层照实返回（它不认识上下文）
            Assert.Equal("—", ChineseFontSizes.Label(double.NaN)); // 界面口径的未知值
            Assert.Null(ChineseFontSizes.NameOf(0));               // 0 不是任何中文字号档位
        }
    }
}
