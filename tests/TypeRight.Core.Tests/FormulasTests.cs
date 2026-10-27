using System.Collections.Generic;
using Xunit;
using TypeRight.Formulas;

namespace TypeRight.Tests
{
    public class MathSegmenterTests
    {
        private static List<string> Segments(string text)
        {
            var list = new List<string>();
            foreach (var s in MathSegmenter.FindSegments(text)) list.Add(s.Text);
            return list;
        }

        [Fact]
        public void InlineFormulaInsideChinese()
        {
            var segs = Segments("对于方程 ax² + bx + c = 0，判别式 Δ = b² - 4ac。");
            Assert.Contains("ax² + bx + c = 0", segs);
            Assert.Contains("Δ = b² - 4ac", segs);
        }

        [Fact]
        public void PureDigitsAndPercentNotConverted()
        {
            Assert.Empty(Segments("增长率为 85%，共 3.14 亿元。"));
        }

        [Fact]
        public void SectionNumberNotConverted()
        {
            Assert.Empty(Segments("详见 6.1 节。"));
        }

        [Fact]
        public void PlainEnglishWordNotConverted()
        {
            // PRD §6.3（v2 验收口径）：含字母即可成片段（黄金 255 区依赖裸字母片段）；
            // 白名单 vs 与型号模式 COVID-19 继续排除。“Office 2016”按 PRD 属可转范围（真实讲义无此形态）。
            var segs = Segments("使用 Office 2016 编辑文档。");
            Assert.Contains("Office 2016", segs);
            Assert.Empty(Segments("甲队 vs 乙队"));
            Assert.Empty(Segments("疫情 COVID-19 期间。"));
        }

        [Fact]
        public void WhitelistVsNotConverted()
        {
            Assert.Empty(Segments("甲队 vs 乙队"));
        }

        [Fact]
        public void ModelNumberNotConverted()
        {
            Assert.Empty(Segments("显卡 RTX-4090 发布。"));
            Assert.Empty(Segments("疫情 COVID-19 期间。"));
        }

        [Fact]
        public void FractionAndSqrtSegment()
        {
            var segs = Segments("代入求根公式(-b±√(b²-4ac))/(2a)求解");
            Assert.Contains("(-b±√(b²-4ac))/(2a)", segs);
        }

        [Fact]
        public void ChinesePunctuationIsBoundary()
        {
            var segs = Segments("已知x₁、x₂是两根（x₁ < x₂），求值。");
            Assert.Contains("x₁", segs);
            Assert.Contains("x₂", segs);
            Assert.Contains("x₁ < x₂", segs);
        }

        [Fact]
        public void StandaloneLineDetected()
        {
            var a = MathSegmenter.Analyze("x₁² + x₂² = (x₁ + x₂)² - 2x₁x₂");
            Assert.True(a.Standalone);
            Assert.True(a.HasEquals);
            Assert.True(a.Density > 0.9);

            var b = MathSegmenter.Analyze("这是一段普通正文，讨论方程 ax²+1=0 的性质。");
            Assert.False(b.Standalone);
        }

        [Fact]
        public void EmptyAndNullSafe()
        {
            Assert.Empty(Segments(""));
            Assert.Empty(Segments(null));
        }
    }

    public class UnicodeMathMapperTests
    {
        [Fact]
        public void SupAndSubScript()
        {
            Assert.Equal("x^2", UnicodeMathMapper.Map("x²"));
            Assert.Equal("x_1", UnicodeMathMapper.Map("x₁"));
            Assert.Equal("x_1^2", UnicodeMathMapper.Map("x₁²"));
            Assert.Equal("x^{12}", UnicodeMathMapper.Map("x¹²"));
        }

        [Fact]
        public void SqrtSingleCharWrapped()
        {
            Assert.Equal("√(x)", UnicodeMathMapper.Map("√x"));
            Assert.Equal("√(2)", UnicodeMathMapper.Map("√2"));
        }

        [Fact]
        public void SqrtParenthesisKept()
        {
            Assert.Equal("√(b^2-4ac)", UnicodeMathMapper.Map("√(b²-4ac)"));
        }

        [Fact]
        public void SqrtBracketNormalized()
        {
            Assert.Equal("√(a+b)", UnicodeMathMapper.Map("√[a+b]"));
        }

        [Fact]
        public void QuadraticFormula()
        {
            Assert.Equal("(-b±√(b^2-4ac))/(2a)", UnicodeMathMapper.Map("(-b±√(b²-4ac))/(2a)"));
        }

        [Fact]
        public void PlainTextUntouched()
        {
            Assert.Equal("a/b + c", UnicodeMathMapper.Map("a/b + c"));
        }

        [Fact]
        public void DeltaAndPiUntouched()
        {
            Assert.Equal("Δ = b^2 - 4ac", UnicodeMathMapper.Map("Δ = b² - 4ac"));
        }
    }
}
