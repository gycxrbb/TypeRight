using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Xunit;
using TypeRight.Formatting;
using TypeRight.Llm;
using TypeRight.Mixing;
using TypeRight.Profiles;
using TypeRight.Roles;

namespace TypeRight.Tests
{
    public class ProfileTests
    {
        [Fact]
        public void BuiltinProfileLoadsAndValidates()
        {
            var p = ProfileStore.BuiltInProfile;
            Assert.Equal("academic-lecture", p.Id);
            Assert.Empty(ProfileValidator.Validate(p));
            Assert.Equal("宋体", p.Roles["body"].EastAsiaFont);
            Assert.Equal("Times New Roman", p.Roles["body"].LatinFont);
            Assert.Equal(1.5, p.Roles["body"].LineSpacing);
            Assert.Equal("Heading 1", p.Roles["heading1"].WordStyle);
        }

        [Fact]
        public void MissingBodyRoleRejected()
        {
            var p = new Profile { Id = "x", Name = "x", SchemaVersion = 1 };
            p.Roles["title"] = new RoleFormat();
            Assert.NotEmpty(ProfileValidator.Validate(p));
        }

        [Fact]
        public void BadColorAndSizeRejected()
        {
            var p = new Profile { Id = "x", Name = "x" };
            p.Roles["body"] = new RoleFormat { Color = "123456", SizePt = -3 };
            var errors = ProfileValidator.Validate(p);
            Assert.Contains(errors, e => e.Contains("color"));
            Assert.Contains(errors, e => e.Contains("sizePt"));
        }

        [Fact]
        public void UnknownRoleKeyRejected()
        {
            var p = new Profile { Id = "x", Name = "x" };
            p.Roles["body"] = new RoleFormat();
            p.Roles["whatever"] = new RoleFormat();
            var errors = ProfileValidator.Validate(p);
            Assert.Contains(errors, e => e.Contains("whatever"));
        }

        [Fact]
        public void JsonRoundTripKeepsCamelCase()
        {
            var p = ProfileStore.BuiltInProfile;
            string json = JsonConvert.SerializeObject(p);
            Assert.Contains("\"schema_version\"", json);
            Assert.Contains("\"eastAsiaFont\"", json);
            var p2 = JsonConvert.DeserializeObject<Profile>(json);
            Assert.Equal(p.Roles.Count, p2.Roles.Count);
            Assert.Equal("黑体", p2.Roles["heading1"].EastAsiaFont);
        }

        [Fact]
        public void RoleKeyParsing()
        {
            Assert.Equal(ParagraphRole.Heading2, ParagraphRoles.TryParse("heading2"));
            Assert.Equal(ParagraphRole.TableText, ParagraphRoles.TryParse("tableText"));
            Assert.Null(ParagraphRoles.TryParse("bogus"));
            Assert.Null(ParagraphRoles.TryParse(""));
            Assert.Equal("tableText", ParagraphRole.TableHeader.FormatKey());
            Assert.Equal("heading1", ParagraphRole.Heading1.FormatKey());
        }

        [Fact]
        public void SaveRejectsInvalidProfileWithoutWritingFile()
        {
            var store = new ProfileStore();
            var bad = new Profile { Id = "bad", Name = "坏方案" };
            bad.Roles["title"] = new RoleFormat { Color = "#ZZZZZZ" }; // 缺 body + 颜色非法
            Assert.Throws<ProfileValidationException>(() => store.Save(bad));
            Assert.False(File.Exists(Path.Combine(ProfileStore.ProfilesDirectory, "bad.json")));
        }
    }

    public class AutoSpacerTests
    {
        [Fact]
        public void InsertSpaceBetweenHanAndLatin()
        {
            Assert.Equal("中文 English 中文", AutoSpacer.Apply("中文English中文", true, PunctuationMode.None));
            Assert.Equal("电压 220V 稳定", AutoSpacer.Apply("电压220V稳定", true, PunctuationMode.None));
        }

        [Fact]
        public void Idempotent()
        {
            string once = AutoSpacer.Apply("中文English混排 220V电压", true, PunctuationMode.None);
            string twice = AutoSpacer.Apply(once, true, PunctuationMode.None);
            Assert.Equal(once, twice);
        }

        [Fact]
        public void NoSpaceInsideEnglishOrBetweenCjkPunctuation()
        {
            Assert.Equal("Hello World", AutoSpacer.Apply("Hello World", true, PunctuationMode.None));
            Assert.Equal("全角。标点，中文", AutoSpacer.Apply("全角。标点，中文", true, PunctuationMode.None));
        }

        [Fact]
        public void FullwidthPunctuationInCjkContext()
        {
            var result = AutoSpacer.Apply("甲说:你好。", false, PunctuationMode.Fullwidth);
            Assert.Equal("甲说：你好。", result);
        }

        [Fact]
        public void DecimalPointProtected()
        {
            Assert.Equal("圆周率是 3.14 左右", AutoSpacer.Apply("圆周率是3.14左右", true, PunctuationMode.Fullwidth));
        }

        [Fact]
        public void FullwidthIdempotent()
        {
            string once = AutoSpacer.Apply("甲说:你好，乙答:再见3.14", false, PunctuationMode.Fullwidth);
            string twice = AutoSpacer.Apply(once, false, PunctuationMode.Fullwidth);
            Assert.Equal(once, twice);
        }
    }

    public class LlmClientTests
    {
        [Fact]
        public void RequestBodyShape()
        {
            var items = new[] { new ParagraphGuessInput { Index = 7, Text = "1.1 背景", FeatureSummary = "len=6;size=12;bold=1" } };
            string body = LlmClient.BuildRequestBody("test-model", items);
            Assert.Contains("\"model\":\"test-model\"", body);
            // user 消息内容是内嵌 JSON 字符串（外层序列化时引号被转义）
            Assert.Contains("\\\"i\\\":7", body);
            Assert.Contains("heading1", body); // system prompt 提到角色集
        }

        [Fact]
        public void ParsePlainJsonResponse()
        {
            var roles = LlmClient.ParseResponse("{\"results\":[{\"i\":1,\"role\":\"heading2\"},{\"i\":2,\"role\":\"body\"}]}");
            Assert.Equal(2, roles.Count);
            Assert.Equal("heading2", roles[0].Role);
        }

        [Fact]
        public void ParseResponseWithProseWrapping()
        {
            var roles = LlmClient.ParseResponse("好的，结果如下：\n```json\n{\"results\":[{\"i\":9,\"role\":\"caption\"}]}\n```");
            Assert.Single(roles);
            Assert.Equal("caption", roles[0].Role);
        }

        [Fact]
        public void ResolverFallsBackOnError()
        {
            var resolver = new AmbiguityResolver(new LlmConfig { Enabled = true, BaseUrl = "http://localhost:1/v1", Model = "m", TimeoutSeconds = 1 });
            var features = new System.Collections.Generic.List<TypeRight.Scanning.ParagraphFeatures>
            {
                new TypeRight.Scanning.ParagraphFeatures { Index = 1, Text = "一段正文" }
            };
            TypeRight.Classification.RoleClassifier.Classify(features);
            var result = resolver.Resolve(features);
            Assert.False(result.UsedLlm);
            Assert.NotNull(result.Error);
            Assert.Equal(ParagraphRole.Body, features[0].Role); // 回落规则结果
        }

        [Fact]
        public void DisabledResolverDoesNothing()
        {
            var resolver = new AmbiguityResolver(new LlmConfig { Enabled = false });
            var result = resolver.Resolve(new System.Collections.Generic.List<TypeRight.Scanning.ParagraphFeatures>());
            Assert.False(result.UsedLlm);
        }
    }

    public class WordFormatUtilsTests
    {
        [Fact]
        public void WdColorRoundTrip()
        {
            // #1A365D → R=0x1A,G=0x36,B=0x5D；WdColor = BGR
            int wd = WordFormatUtils.RgbToWdColor(0x1A365D);
            Assert.Equal(0x1A, wd & 0xFF);         // R 在低位
            Assert.Equal(0x36, (wd >> 8) & 0xFF);  // G
            Assert.Equal(0x5D, (wd >> 16) & 0xFF); // B
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

        [Fact]
        public void BuiltinStyleMapping()
        {
            Assert.Equal(NetOffice.WordApi.Enums.WdBuiltinStyle.wdStyleHeading1, WordFormatUtils.TryGetBuiltinStyle("Heading 1"));
            Assert.Equal(NetOffice.WordApi.Enums.WdBuiltinStyle.wdStyleTitle, WordFormatUtils.TryGetBuiltinStyle("title"));
            Assert.Null(WordFormatUtils.TryGetBuiltinStyle("My Custom Style"));
        }
    }
}
