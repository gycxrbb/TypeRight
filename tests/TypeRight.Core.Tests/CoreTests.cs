using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Xunit;
using TypeRight.Formatting;
using TypeRight.Profiles;
using TypeRight.Signature;

namespace TypeRight.Tests
{
    public class FormatSignatureTests
    {
        private static FormatSignature Sig(double size = 10.5, bool bold = false, int color = -1,
            string east = "宋体", string latin = "Times New Roman", string align = "l", int italic = 0)
        {
            return new FormatSignature
            {
                EastAsiaFont = east, LatinFont = latin, SizePt = size,
                BoldRaw = bold ? -1 : 0, ItalicRaw = italic, ColorRgb = color, Alignment = align
            };
        }

        [Fact]
        public void StrictEqualityAllDims()
        {
            Assert.True(Sig().Equals(Sig()));
            Assert.False(Sig().Equals(Sig(size: 12)));
            Assert.False(Sig().Equals(Sig(bold: true)));
            Assert.False(Sig().Equals(Sig(color: 0xFF0000)));
            Assert.False(Sig().Equals(Sig(east: "黑体")));
            Assert.False(Sig().Equals(Sig(align: "c")));
            // 字体名大小写不敏感
            Assert.True(Sig(east: "simsun").Equals(Sig(east: "SimSun")));
        }

        [Fact]
        public void NanSizeEqualsNan()
        {
            var a = Sig(); a.SizePt = double.NaN;
            var b = Sig(); b.SizePt = double.NaN;
            Assert.True(a.Equals(b));
            Assert.False(a.Equals(Sig()));
        }

        [Fact]
        public void LooseIgnoresSelectedDims()
        {
            var sample = Sig(size: 22, bold: true, color: 0x1A365D);
            var module = Sig(size: 15, bold: true, color: 0x1A365D);
            Assert.False(sample.Matches(module, FormatSignature.IgnoreDims.None), "严格模式字号不等");
            Assert.True(sample.Matches(module, FormatSignature.IgnoreDims.SizePt), "忽略字号后命中");

            var colored = Sig(size: 22, bold: true, color: 0xFF0000);
            Assert.False(sample.Matches(colored, FormatSignature.IgnoreDims.SizePt), "颜色仍参与比较");
            Assert.True(sample.Matches(colored, FormatSignature.IgnoreDims.SizePt | FormatSignature.IgnoreDims.Color));

            var notBold = Sig(size: 22, bold: false, color: 0x1A365D);
            Assert.True(sample.Matches(notBold, FormatSignature.IgnoreDims.Bold));

            // 字体与对齐永远参与比较
            Assert.False(sample.Matches(Sig(size: 22, bold: true, color: 0x1A365D, east: "黑体"), FormatSignature.IgnoreDims.SizePt | FormatSignature.IgnoreDims.Color | FormatSignature.IgnoreDims.Bold));
        }

        [Fact]
        public void MatchSpecNullDimMeansIgnore()
        {
            var spec = new MatchSpec { EastAsiaFont = "宋体" };  // 其余维度 null = 忽略
            Assert.True(spec.MatchesPara(Sig(size: 22, bold: true, color: 0x1A365D)));
            Assert.True(spec.MatchesPara(Sig(size: 9)));
            Assert.False(spec.MatchesPara(Sig(east: "黑体")));

            var full = new MatchSpec { EastAsiaFont = "宋体", SizePt = 12.5, Bold = true, Color = "#2B6CB0" };
            Assert.True(full.MatchesPara(Sig(size: 12.5, bold: true, color: 0x2B6CB0)));
            Assert.False(full.MatchesPara(Sig(size: 12.5, bold: true, color: 0x2B6CB1)));
            Assert.False(full.MatchesPara(Sig(size: 12.6, bold: true, color: 0x2B6CB0)));

            // ignore 列表让维度失效
            full.Ignore = new List<string> { "color" };
            Assert.True(full.MatchesPara(Sig(size: 12.5, bold: true, color: 0x112233)));

            var catchAll = new MatchSpec { CatchAll = true };
            Assert.True(catchAll.MatchesPara(Sig()));
        }
    }

    public class ProfileV2Tests
    {
        private static ProfileV2 Valid()
        {
            return new ProfileV2
            {
                Id = "math-lecture", Name = "数学讲义",
                Rules = new List<ProfileRule>
                {
                    new ProfileRule { Order = 1,
                        Match = new MatchSpec { SizePt = 15, Bold = true, Color = "#1A365D", Ignore = new List<string> { "sizePt" } },
                        Apply = new ApplySpec { EastAsiaFont = "楷体", LatinFont = "Times New Roman" } },
                    new ProfileRule { Order = 2,
                        Match = new MatchSpec { CatchAll = true },
                        Apply = new ApplySpec { EastAsiaFont = "宋体", LatinFont = "Times New Roman" } }
                }
            };
        }

        [Fact]
        public void ValidProfilePasses()
        {
            Assert.Empty(ProfileV2Validator.Validate(Valid()));
        }

        [Fact]
        public void FieldLevelErrors()
        {
            var p = Valid();
            p.SchemaVersion = 1;
            p.Rules[0].Order = 5;
            p.Rules.Add(new ProfileRule { Order = 3,
                Match = new MatchSpec { Color = "ZZZ" }, Apply = new ApplySpec { SizePt = -1 } });
            var errors = ProfileV2Validator.Validate(p);

            Assert.Contains(errors, e => e.StartsWith("schema_version:"));
            Assert.Contains(errors, e => e.StartsWith("rules[0].order:"));
            Assert.Contains(errors, e => e.Contains("match.color"));
            Assert.Contains(errors, e => e.Contains("apply.sizePt"));
            Assert.Contains(errors, e => e.Contains("catchAll") && e.Contains("末位"));
        }

        [Fact]
        public void EmptyMatchAndApplyRejected()
        {
            var p = Valid();
            p.Rules[0].Match = new MatchSpec();           // 无任何维度
            p.Rules[1].Apply = new ApplySpec();           // 全 null
            var errors = ProfileV2Validator.Validate(p);
            Assert.Contains(errors, e => e.Contains("至少要有一个签名维度"));
            Assert.Contains(errors, e => e.Contains("不会改动任何内容"));
        }

        [Fact]
        public void JsonRoundTripKeepsNullsAndOrder()
        {
            ProfileV2 p = Valid();
            p.Rules[0].Apply.SizePt = null;   // null = 保持源值，必须可序列化往返
            string json = JsonConvert.SerializeObject(p);
            Assert.Contains("\"schema_version\":2", json);

            var back = JsonConvert.DeserializeObject<ProfileV2>(json);
            Assert.Equal(2, back.Rules.Count);
            Assert.Null(back.Rules[0].Apply.SizePt);
            Assert.Null(back.Rules[0].Apply.Bold);
            Assert.Equal("楷体", back.Rules[0].Apply.EastAsiaFont);
            Assert.True(back.Rules[1].Match.CatchAll == true);
            // ignore 含 sizePt，且仅含 sizePt
            Assert.True((back.Rules[0].Match.IgnoreValue & FormatSignature.IgnoreDims.SizePt) != 0);
            Assert.True((back.Rules[0].Match.IgnoreValue & ~(FormatSignature.IgnoreDims.SizePt | FormatSignature.IgnoreDims.Color | FormatSignature.IgnoreDims.Bold)) == 0);
            Assert.Empty(ProfileV2Validator.Validate(back));
        }

        [Fact]
        public void SaveRejectsInvalid()
        {
            var store = new ProfileStoreV2();
            var bad = Valid();
            bad.Id = "";
            Assert.Throws<ProfileValidationException>(() => store.Save(bad));
        }
    }

    public class WordFormatUtilsTests
    {
        [Fact]
        public void WdColorRoundTrip()
        {
            int wd = WordFormatUtils.RgbToWdColor(0x1A365D);
            Assert.Equal(0x1A, wd & 0xFF);
            Assert.Equal(0x36, (wd >> 8) & 0xFF);
            Assert.Equal(0x5D, (wd >> 16) & 0xFF);
            Assert.Equal(0x1A365D, WordFormatUtils.WdColorToRgb(wd));
        }

        [Fact]
        public void HexParseAndInvalid()
        {
            Assert.Equal(0xFF0000, WordFormatUtils.ParseHexColor("#FF0000"));
            Assert.Null(WordFormatUtils.ParseHexColor("FF0000"));
            Assert.Null(WordFormatUtils.ParseHexColor("#GGHHII"));
            Assert.Null(WordFormatUtils.ParseHexColor(null));
        }

        [Fact]
        public void UndefinedSentinels()
        {
            Assert.True(double.IsNaN(WordFormatUtils.NormalizeSize(9999999f)));
            Assert.Equal(10.5, WordFormatUtils.NormalizeSize(10.5f));
        }
    }
}
