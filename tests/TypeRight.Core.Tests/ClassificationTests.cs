using System.Collections.Generic;
using System.Linq;
using Xunit;
using TypeRight.Classification;
using TypeRight.Roles;
using TypeRight.Scanning;

namespace TypeRight.Tests
{
    public class RuleTableTests
    {
        private static ParagraphFeatures F(string text)
        {
            return new ParagraphFeatures { Text = text, IsEmpty = text.Trim().Length == 0, Alignment = "l" };
        }

        private static void ApplyPatterns(params ParagraphFeatures[] fs)
        {
            foreach (var f in fs) f.NumberingPattern = DocumentScanner.DetectNumberingPattern(f.Text);
        }

        [Fact]
        public void R1_EmptyParaSkipped()
        {
            ApplyPatterns(F(""));
            var f = F("");
            var ctx = new ClassifierContext();
            Assert.Equal(ParagraphRole.Skip, RuleTable.EvaluateHighConfidence(f, ctx));
        }

        [Fact]
        public void R2_TableRules()
        {
            var ctx = new ClassifierContext();
            var header = new ParagraphFeatures { Text = "参数说明", BoldRaw = -1, InTable = true, IsFirstTableRow = true };
            var cell = new ParagraphFeatures { Text = "普通单元格", InTable = true, IsFirstTableRow = false };
            var firstRowNotBold = new ParagraphFeatures { Text = "无表头", InTable = true, IsFirstTableRow = true, BoldRaw = 0 };
            Assert.Equal(ParagraphRole.TableHeader, RuleTable.EvaluateHighConfidence(header, ctx));
            Assert.Equal(ParagraphRole.TableText, RuleTable.EvaluateHighConfidence(cell, ctx));
            Assert.Equal(ParagraphRole.TableText, RuleTable.EvaluateHighConfidence(firstRowNotBold, ctx));
        }

        [Fact]
        public void R3_ExistingHeadingStyleTrusted()
        {
            var ctx = new ClassifierContext();
            var f = new ParagraphFeatures { Text = "奇怪的短句没有编号", HeadingStyleLevel = 2 };
            Assert.Equal(ParagraphRole.Heading2, RuleTable.EvaluateHighConfidence(f, ctx));
        }

        [Fact]
        public void R4_TitleFirstNonEmptyBigBold()
        {
            var ctx = new ClassifierContext { BodyFontSize = 10.5 };
            var f = new ParagraphFeatures { Text = "课程讲义标题", IsFirstNonEmpty = true, FontSize = 20, BoldRaw = -1 };
            Assert.Equal(ParagraphRole.Title, RuleTable.EvaluateHighConfidence(f, ctx));

            var tooSmall = new ParagraphFeatures { Text = "小标题", IsFirstNonEmpty = true, FontSize = 12, BoldRaw = -1 };
            Assert.Null(RuleTable.EvaluateHighConfidence(tooSmall, ctx));

            var notBold = new ParagraphFeatures { Text = "不粗的标题", IsFirstNonEmpty = true, FontSize = 20, BoldRaw = 0 };
            Assert.Null(RuleTable.EvaluateHighConfidence(notBold, ctx));
        }

        [Fact]
        public void R5_ChapterPattern()
        {
            var ctx = new ClassifierContext();
            Assert.Equal(ParagraphRole.Heading1, RuleTable.EvaluateHighConfidence(Mk("第一章 绪论"), ctx));
            Assert.Equal(ParagraphRole.Heading1, RuleTable.EvaluateHighConfidence(Mk("第3章 方程"), ctx));
            Assert.Equal(ParagraphRole.Heading1, RuleTable.EvaluateHighConfidence(Mk("模块二：例题精讲"), ctx));
            // 过长不判
            var longChapter = Mk("第一章 " + new string('长', 60));
            Assert.Null(RuleTable.EvaluateHighConfidence(longChapter, ctx));
        }

        [Fact]
        public void R6_R7_NumericPatterns()
        {
            var ctx = new ClassifierContext();
            Assert.Equal(ParagraphRole.Heading3, RuleTable.EvaluateHighConfidence(Mk("1.1.2 根的判别"), ctx));
            Assert.Equal(ParagraphRole.Heading2, RuleTable.EvaluateHighConfidence(Mk("1.1 一元二次方程"), ctx));
        }

        [Fact]
        public void R8_CaptionPattern()
        {
            var ctx = new ClassifierContext();
            Assert.Equal(ParagraphRole.Caption, RuleTable.EvaluateHighConfidence(Mk("图 1 函数图像"), ctx));
            Assert.Equal(ParagraphRole.Caption, RuleTable.EvaluateHighConfidence(Mk("表 2 参数列表"), ctx));
            Assert.Equal(ParagraphRole.Caption, RuleTable.EvaluateHighConfidence(Mk("Figure 3 Result"), ctx));
        }

        [Fact]
        public void R9_AutoNumberingListItem()
        {
            var ctx = new ClassifierContext();
            var f = new ParagraphFeatures { Text = "自动编号的列表项", HasAutoNumbering = true };
            Assert.Equal(ParagraphRole.ListItem, RuleTable.EvaluateHighConfidence(f, ctx));
        }

        [Fact]
        public void R10_FormulaLine()
        {
            var ctx = new ClassifierContext();
            var f = new ParagraphFeatures
            {
                Text = "x₁² + x₂² = (x₁ + x₂)² - 2x₁x₂",
                StandaloneMath = true,
                MathDensity = 1.0,
                HasEquals = true
            };
            Assert.Equal(ParagraphRole.FormulaLine, RuleTable.EvaluateHighConfidence(f, ctx));
        }

        private static ParagraphFeatures Mk(string text)
        {
            var f = F(text);
            f.NumberingPattern = DocumentScanner.DetectNumberingPattern(text);
            f.CaptionPattern = System.Text.RegularExpressions.Regex.IsMatch(text,
                @"^(图|表|Figure|Table|【图|【表)\s*\S", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            return f;
        }
    }

    public class ClassifierFlowTests
    {
        [Fact]
        public void ClusterFallbackPromotesBoldClustersToHeadings()
        {
            var features = new List<ParagraphFeatures>();
            // 未命中任何模式的段落：加粗 15pt 簇 ×2、加粗 13pt 簇 ×1、正文 10.5pt 簇 ×10
            // 注意：首个 15pt 加粗段按 PRD R4 判为主标题（字号≥正文+4 且加粗的第一个非空段）
            features.Add(Mk("先讲讲历史背景与课程要求", 15, true));
            features.Add(Mk("再讲讲发展脉络和研究方法", 15, true));
            features.Add(Mk("细分的知识点补充说明", 13, true));
            for (int i = 0; i < 10; i++)
                features.Add(Mk("这是正文段落第 " + i + " 行，长度足够长以避免其他模式命中。", 10.5, false));

            RoleClassifier.Classify(features);

            Assert.Equal(ParagraphRole.Title, features[0].Role);      // R4 高置信
            Assert.True(features[0].HighConfidence);
            Assert.Equal(ParagraphRole.Heading1, features[1].Role);   // R11 聚类补位
            Assert.False(features[1].HighConfidence);
            Assert.Equal(ParagraphRole.Heading2, features[2].Role);
            Assert.All(features.Skip(3), f => Assert.Equal(ParagraphRole.Body, f.Role));
        }

        [Fact]
        public void BodyFontSizeIsModeOfLongParagraphs()
        {
            var features = new List<ParagraphFeatures>();
            for (int i = 0; i < 5; i++) features.Add(Mk("这是一段比较长的正文内容用来估计正文字号的众数。", 10.5, false));
            features.Add(Mk("短小段落", 12, false));
            Assert.Equal(10.5, RoleClassifier.ComputeBodyFontSize(features));
        }

        private static ParagraphFeatures Mk(string text, double size, bool bold)
        {
            var f = new ParagraphFeatures
            {
                Text = text,
                IsEmpty = false,
                FontSize = size,
                BoldRaw = bold ? -1 : 0,
                Alignment = "l",
                NumberingPattern = DocumentScanner.DetectNumberingPattern(text)
            };
            return f;
        }
    }
}
