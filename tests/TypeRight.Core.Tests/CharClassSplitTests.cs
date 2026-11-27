using System.Collections.Generic;
using TypeRight.Signature;
using Xunit;

namespace TypeRight.Tests
{
    /// <summary>按字符类型切分/过滤（同类选择"汉字归汉字、数字归数字"的底层）。
    ///
    /// 回归背景：FilterRuns 第一版遇到异类片段只置 has=false 而没有 flush，
    /// 导致前面已累积的片段全部丢失（真机实测：177 字段落只保留 1 个片段，
    /// 过滤后什么都没改）。这里逐项锁死。</summary>
    public class CharClassSplitTests
    {
        [Fact]
        public void SplitByClass_SeparatesMixedContent()
        {
            // "商雨萤同志，女，2005年" 按**连续同类型**切分：
            //   商雨萤同志(5,Cjk) | 女(1,Cjk) | 2005(4,Digit) | 年(1,Cjk)
            // 逗号不产生片段、且会打断连续性，所以两段中文不会合并 → 共 4 段。
            // （合并只在 FilterRuns 里按需做，SplitByClass 保持原始切分。）
            var runs = CharClassDetector.SplitByClass("商雨萤同志，女，2005年");
            Assert.Equal(4, runs.Count);
            Assert.Equal(CharClass.Cjk, runs[0].Class);
            Assert.Equal(5, runs[0].Length);        // 商雨萤同志
            Assert.Equal(CharClass.Cjk, runs[1].Class);
            Assert.Equal(1, runs[1].Length);        // 女（逗号被跳过）
            Assert.Equal(CharClass.Digit, runs[2].Class);
            Assert.Equal(4, runs[2].Length);        // 2005
            Assert.Equal(CharClass.Cjk, runs[3].Class);
            Assert.Equal(1, runs[3].Length);        // 年
        }

        /// <summary>FilterRuns 应把被标点隔开的同类型片段合并（少刷一次格式）。</summary>
        [Fact]
        public void FilterRuns_MergesAcrossPunctuation()
        {
            var runs = CharClassDetector.SplitByClass("商雨萤同志，女");
            var cjk = CharClassDetector.FilterRuns(runs, CharClass.Cjk, mergeAcrossPunctuation: true);

            Assert.Single(cjk);
            Assert.Equal(CharClass.Cjk, cjk[0].Class);
            // 跨过中间那个逗号后，两段中文连成 0..6（商雨萤同志，女）
            Assert.Equal(7, cjk[0].Length);
        }

        /// <summary>不合并时，标点两侧各成一段 —— 这是 FilterRuns 未开合并的语义。</summary>
        [Fact]
        public void FilterRuns_NoMergeKeepsSeparate()
        {
            var runs = CharClassDetector.SplitByClass("商雨萤同志，女");
            var cjk = CharClassDetector.FilterRuns(runs, CharClass.Cjk, mergeAcrossPunctuation: false);

            Assert.Equal(2, cjk.Count);
            Assert.Equal(5, cjk[0].Length);
            Assert.Equal(1, cjk[1].Length);
        }

        /// <summary>回归坑：FilterRuns 遇到异类必须 flush，否则前面片段全丢。</summary>
        [Fact]
        public void FilterRuns_DoesNotDropEarlierSegments()
        {
            var runs = CharClassDetector.SplitByClass("商雨萤2005年");
            var cjk = CharClassDetector.FilterRuns(runs, CharClass.Cjk, mergeAcrossPunctuation: false);

            // 商雨萤 + 年，两段中文被数字隔开
            Assert.Equal(2, cjk.Count);
            Assert.Equal("商雨萤", "商雨萤2005年".Substring(cjk[0].Start, cjk[0].Length));
            Assert.Equal("年", "商雨萤2005年".Substring(cjk[1].Start, cjk[1].Length));
        }

        [Fact]
        public void SplitByClass_SkipsPunctuationWithoutProducingRuns()
        {
            var runs = CharClassDetector.SplitByClass("，。！？");
            Assert.Empty(runs);
        }

        [Fact]
        public void SplitByClass_EmptyInput_ReturnsEmpty()
        {
            Assert.Empty(CharClassDetector.SplitByClass(""));
            Assert.Empty(CharClassDetector.SplitByClass(null));
        }

        [Fact]
        public void FilterRuns_KeepsAllMatchingRuns_NotJustLast()
        {
            // 核心回归：三个中文片段被数字隔开，必须全部保留
            var runs = CharClassDetector.SplitByClass("甲乙1丙丁2戊");
            var keep = CharClassDetector.FilterRuns(runs, CharClass.Cjk, mergeAcrossPunctuation: false);
            Assert.Equal(3, keep.Count);
            Assert.Equal(0, keep[0].Start);
            Assert.Equal(2, keep[0].Length);        // 甲乙
        }

        [Fact]
        public void FilterRuns_DigitOnly_KeepsDigits()
        {
            var runs = CharClassDetector.SplitByClass("第1题2个3次");
            var keep = CharClassDetector.FilterRuns(runs, CharClass.Digit, mergeAcrossPunctuation: false);
            Assert.Equal(3, keep.Count);
            Assert.Equal(CharClass.Digit, keep[0].Class);
        }

        [Fact]
        public void FilterRuns_MergesAcrossSinglePunctuation()
        {
            // "商雨萤同志，女" —— 逗号隔开，合并后应是一段连续中文
            var runs = CharClassDetector.SplitByClass("商雨萤同志，女");
            var merged = CharClassDetector.FilterRuns(runs, CharClass.Cjk, mergeAcrossPunctuation: true);
            Assert.Single(merged);
            Assert.Equal(0, merged[0].Start);
            Assert.Equal(7, merged[0].Length);      // 含逗号，整段 7 字
        }

        [Fact]
        public void FilterRuns_DoesNotMergeAcrossLongGaps()
        {
            // 两个中文片段中间隔了 "12345"（5 字），不应合并
            var runs = CharClassDetector.SplitByClass("甲乙12345丙丁");
            var keep = CharClassDetector.FilterRuns(runs, CharClass.Cjk, mergeAcrossPunctuation: true);
            Assert.Equal(2, keep.Count);
        }

        [Fact]
        public void FilterRuns_RealParagraph_ProducesManySegments()
        {
            // 真实正文（177 字的简化版）：中文片段必须远多于 1 个
            string text = "商雨萤同志，女，2005年5月出生，2023级数据科学与大数据技术专业本科生。";
            var runs = CharClassDetector.SplitByClass(text);
            var keep = CharClassDetector.FilterRuns(runs, CharClass.Cjk, mergeAcrossPunctuation: true);
            Assert.True(keep.Count >= 4, "中文片段数应 >= 4，实际 " + keep.Count);
        }

        [Fact]
        public void FilterRuns_NeverReturnsOverlappingOrOutOfRange()
        {
            string text = "预备期内，党支部按照要求，共召开会议12次，组织活动3场。";
            var runs = CharClassDetector.SplitByClass(text);
            foreach (CharClass want in new[] { CharClass.Cjk, CharClass.Digit, CharClass.Latin })
            {
                var keep = CharClassDetector.FilterRuns(runs, want, mergeAcrossPunctuation: true);
                int lastEnd = -1;
                foreach (var r in keep)
                {
                    Assert.True(r.Start >= 0, "Start 越界");
                    Assert.True(r.End <= text.Length, "End 越界: " + r.End + " > " + text.Length);
                    Assert.True(r.Length > 0, "空片段");
                    Assert.True(r.Start >= lastEnd, "片段重叠: " + r.Start + " < " + lastEnd);
                    lastEnd = r.End;
                }
            }
        }

        [Theory]
        [InlineData('中', CharClass.Cjk)]
        [InlineData('1', CharClass.Digit)]
        [InlineData('a', CharClass.Latin)]
        [InlineData('，', CharClass.None)]
        [InlineData(' ', CharClass.None)]
        public void ClassOf_SingleChar(char c, CharClass expected)
        {
            Assert.Equal(expected, CharClassDetector.ClassOf(c));
        }
    }
}
