using System.Collections.Generic;
using System.Linq;
using TypeRight.Profiles;
using TypeRight.Signature;
using Xunit;

namespace TypeRight.Tests
{
    /// <summary>规则归一化：同一次会话连着改多个属性，必须合并成一条规则。</summary>
    public class OpNormalizerTests
    {
        private static MatchSpec Match16()
        {
            return new MatchSpec
            {
                EastAsiaFont = "仿宋",
                LatinFont = "仿宋",
                SizePt = 16,
                Bold = false,
                Color = "#000000"
            };
        }

        [Fact]
        public void SameMatch_SameSpec_IsTrue()
        {
            Assert.True(OpNormalizer.SameMatch(Match16(), Match16()));
        }

        [Fact]
        public void SameMatch_DifferentFont_IsFalse()
        {
            var b = Match16();
            b.EastAsiaFont = "宋体";
            Assert.False(OpNormalizer.SameMatch(Match16(), b));
        }

        [Fact]
        public void SameMatch_DifferentIgnore_IsFalse()
        {
            var a = Match16();
            a.Ignore = new List<string> { "sizePt" };
            var b = Match16();
            Assert.False(OpNormalizer.SameMatch(a, b));
        }

        [Fact]
        public void SameMatch_DifferentTextDomain_IsFalse()
        {
            var a = Match16(); a.InTable = false;
            var b = Match16(); b.InTable = true;
            Assert.False(OpNormalizer.SameMatch(a, b));
        }

        /// <summary>这就是线上那个 bug：一次会话改字体→字号→字号，产出 3 条同样 match 的规则。</summary>
        [Fact]
        public void Normalize_MergesConsecutiveSameMatchRules()
        {
            var raw = new List<ProfileRule>
            {
                new ProfileRule { Order = 1, Match = Match16(),
                    Apply = new ApplySpec { EastAsiaFont = "仿宋", LatinFont = "Times New Roman" } },
                new ProfileRule { Order = 2, Match = Match16(),
                    Apply = new ApplySpec { SizePt = 15 } },
                new ProfileRule { Order = 3, Match = Match16(),
                    Apply = new ApplySpec { SizePt = 14 } },
            };

            var merged = OpNormalizer.Normalize(raw);

            Assert.Single(merged);
            Assert.Equal(1, merged[0].Order);
            Assert.Equal("仿宋", merged[0].Apply.EastAsiaFont);
            Assert.Equal("Times New Roman", merged[0].Apply.LatinFont);
            Assert.Equal(14, merged[0].Apply.SizePt);   // last-write-wins：最终是 14
        }

        [Fact]
        public void Normalize_MergeKeepsLaterNonNull()
        {
            var earlier = new ApplySpec { EastAsiaFont = "仿宋", SizePt = 16 };
            var later = new ApplySpec { Bold = true };
            var m = OpNormalizer.MergeApply(earlier, later);

            Assert.Equal("仿宋", m.EastAsiaFont);   // earlier 独有的保留
            Assert.Equal(16, m.SizePt);
            Assert.True(m.Bold == true);            // later 的新值生效
        }

        [Fact]
        public void Normalize_DifferentMatch_StaysSeparate()
        {
            var other = Match16();
            other.EastAsiaFont = "黑体";

            var raw = new List<ProfileRule>
            {
                new ProfileRule { Order = 1, Match = Match16(), Apply = new ApplySpec { EastAsiaFont = "仿宋" } },
                new ProfileRule { Order = 2, Match = other,    Apply = new ApplySpec { EastAsiaFont = "黑体" } },
            };

            var merged = OpNormalizer.Normalize(raw);
            Assert.Equal(2, merged.Count);
            Assert.Equal(1, merged[0].Order);
            Assert.Equal(2, merged[1].Order);
        }

        /// <summary>兜底规则绝不能与普通规则合并——合并会把"其余段落"的语义抹掉。</summary>
        [Fact]
        public void Normalize_CatchAllNeverMerges()
        {
            var raw = new List<ProfileRule>
            {
                new ProfileRule { Order = 1, Match = Match16(), Apply = new ApplySpec { EastAsiaFont = "仿宋" } },
                new ProfileRule { Order = 2, Match = new MatchSpec { CatchAll = true }, Apply = new ApplySpec { EastAsiaFont = "宋体" } },
                new ProfileRule { Order = 3, Match = new MatchSpec { CatchAll = true }, Apply = new ApplySpec { SizePt = 12 } },
            };

            var merged = OpNormalizer.Normalize(raw);
            Assert.Equal(3, merged.Count);
        }

        [Fact]
        public void Normalize_ReordersContiguously()
        {
            var other = Match16(); other.EastAsiaFont = "黑体";
            var raw = new List<ProfileRule>
            {
                new ProfileRule { Order = 7, Match = Match16(), Apply = new ApplySpec { EastAsiaFont = "仿宋" } },
                new ProfileRule { Order = 9, Match = other, Apply = new ApplySpec { EastAsiaFont = "黑体" } },
            };
            var merged = OpNormalizer.Normalize(raw);
            Assert.Equal(1, merged[0].Order);
            Assert.Equal(2, merged[1].Order);
        }
    }

    /// <summary>方案 → 人话。用户看不懂"规则 3 条 / True"，这里锁定可读文案口径。</summary>
    public class ProfileDescriberTests
    {
        private static ProfileV2 Profile(params ProfileRule[] rules)
        {
            var p = new ProfileV2 { Id = "x", Name = "预备党员转正" };
            p.Rules = rules.ToList();
            return p;
        }

        [Fact]
        public void Describe_RuleCountIsHuman()
        {
            ProfileV2 p = Profile(
                new ProfileRule { Order = 1, Match = new MatchSpec { EastAsiaFont = "仿宋" }, Apply = new ApplySpec { SizePt = 14 } },
                new ProfileRule { Order = 2, Match = new MatchSpec { EastAsiaFont = "黑体" }, Apply = new ApplySpec { SizePt = 16 } });

            ProfileSummary s = ProfileDescriber.Describe(p);
            Assert.Equal("改了 2 处格式", s.RuleCountText);
        }

        [Fact]
        public void Describe_SingleChangeReadsNaturally()
        {
            ProfileV2 p = Profile(
                new ProfileRule { Order = 1, Match = new MatchSpec { EastAsiaFont = "仿宋" }, Apply = new ApplySpec { SizePt = 14 } });

            Assert.Equal("改了 1 处格式", ProfileDescriber.Describe(p).RuleCountText);
        }

        /// <summary>这就是用户描述的那句话："把仿宋 黑色 四号 无加粗 的字号变成三号"。</summary>
        [Fact]
        public void DescribeChange_ReportsFromToForSize()
        {
            var match = new MatchSpec
            {
                EastAsiaFont = "仿宋", Color = "#000000", SizePt = 14, Bold = false
            };
            var apply = new ApplySpec { SizePt = 16 };   // 只改字号

            string change = ProfileDescriber.DescribeChange(match, apply);

            Assert.Equal("字号 四号 → 三号", change);
        }

        [Fact]
        public void DescribeChange_OnlyMentionsWhatChanged()
        {
            var match = new MatchSpec { EastAsiaFont = "仿宋", SizePt = 14 };
            var apply = new ApplySpec { SizePt = 16 };   // 字体没动

            string change = ProfileDescriber.DescribeChange(match, apply);
            Assert.DoesNotContain("仿宋", change);       // 没改的字体不该出现在"改了什么"里
            Assert.Contains("字号", change);
        }

        [Fact]
        public void DescribeChange_NewSizeWithoutMatchSize()
        {
            var change = ProfileDescriber.DescribeChange(new MatchSpec { EastAsiaFont = "仿宋" },
                new ApplySpec { SizePt = 16 });
            Assert.Equal("字号 → 三号", change);
        }

        [Fact]
        public void DescribeChange_FontChangeShowsFromTo()
        {
            var change = ProfileDescriber.DescribeChange(
                new MatchSpec { EastAsiaFont = "宋体" }, new ApplySpec { EastAsiaFont = "仿宋" });
            Assert.Contains("宋体", change);
            Assert.Contains("仿宋", change);
            Assert.Contains("→", change);
        }

        [Fact]
        public void DescribeChange_BoldToggles()
        {
            var change = ProfileDescriber.DescribeChange(
                new MatchSpec { Bold = false }, new ApplySpec { Bold = true });
            Assert.Contains("→", change);
            Assert.Contains("加粗", change);
        }

        [Fact]
        public void DescribeChange_NothingChangedReturnsNull()
        {
            Assert.Null(ProfileDescriber.DescribeChange(new MatchSpec { EastAsiaFont = "仿宋" }, new ApplySpec()));
        }

        /// <summary>字号必须走中文字号口径，不能吐裸数字 16。</summary>
        [Fact]
        public void DescribeMatch_UsesChineseSizeName()
        {
            var m = new MatchSpec { EastAsiaFont = "仿宋", SizePt = 16, Bold = false };
            string text = ProfileDescriber.DescribeMatch(m);

            Assert.Contains("仿宋", text);
            Assert.Contains("三号", text);        // 16 磅 = 三号，不是裸数字
            Assert.Contains("不加粗", text);
        }

        [Fact]
        public void DescribeMatch_NonStandardSizeFallsBackToNumber()
        {
            // 12.5 磅不在中文字号表里 → 显示数值，不硬套中文名
            string text = ProfileDescriber.DescribeMatch(new MatchSpec { SizePt = 12.5 });
            Assert.Contains("12.5 磅", text);
        }

        /// <summary>匹配条件的读序要与用户口述一致：字体 → 颜色 → 字号 → 加粗。</summary>
        [Fact]
        public void DescribeMatch_OrderMatchesHowUsersSayIt()
        {
            var m = new MatchSpec { EastAsiaFont = "仿宋", Color = "#000000", SizePt = 14, Bold = false };
            string text = ProfileDescriber.DescribeMatch(m);
            Assert.StartsWith("仿宋", text);
            Assert.True(text.IndexOf("黑色") < text.IndexOf("四号"), "颜色应在字号前");
            Assert.True(text.IndexOf("四号") < text.IndexOf("不加粗"), "字号应在加粗前");
        }

        [Fact]
        public void DescribeMatch_ShowsIgnoredDims()
        {
            var m = new MatchSpec { EastAsiaFont = "仿宋", Ignore = new List<string> { "sizePt", "color" } };
            string text = ProfileDescriber.DescribeMatch(m);
            Assert.Contains("不看", text);
            Assert.Contains("字号", text);
            Assert.Contains("颜色", text);
        }

        [Fact]
        public void DescribeApply_LatinFontIsCalledOut()
        {
            var a = new ApplySpec { EastAsiaFont = "仿宋", SizePt = 14, LatinFont = "Times New Roman" };
            string text = ProfileDescriber.DescribeApply(a);

            Assert.Contains("仿宋", text);
            Assert.Contains("四号", text);
            Assert.Contains("西文 Times New Roman", text);
        }

        [Fact]
        public void DescribeApply_EmptyMeansNoChange()
        {
            Assert.Equal("不改变", ProfileDescriber.DescribeApply(new ApplySpec()));
        }

        [Fact]
        public void FormulaNote_OnlyWhenEnabled()
        {
            ProfileV2 on = Profile(new ProfileRule { Order = 1, Match = new MatchSpec { EastAsiaFont = "仿宋" }, Apply = new ApplySpec { SizePt = 14 } });
            on.ConvertFormulas = true;
            Assert.NotNull(ProfileDescriber.Describe(on).FormulaNote);

            ProfileV2 off = Profile(new ProfileRule { Order = 1, Match = new MatchSpec { EastAsiaFont = "仿宋" }, Apply = new ApplySpec { SizePt = 14 } });
            off.ConvertFormulas = false;
            Assert.Null(ProfileDescriber.Describe(off).FormulaNote);
        }

        [Fact]
        public void ColorLabel_KnownColors()
        {
            Assert.Equal("黑色", ProfileDescriber.ColorLabel("#000000"));
            Assert.Equal("红色", ProfileDescriber.ColorLabel("#FF0000"));
            Assert.Equal("#1A365D", ProfileDescriber.ColorLabel("#1A365D"));   // 不认识的不得吞掉
        }

        [Fact]
        public void IsBroad_DetectsWildcardMatch()
        {
            Assert.True(ProfileDescriber.IsBroad(new MatchSpec()));
            Assert.False(ProfileDescriber.IsBroad(new MatchSpec { EastAsiaFont = "仿宋" }));
        }
    }
}
