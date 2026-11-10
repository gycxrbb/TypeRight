using System.Collections.Generic;
using TypeRight.Sequences;
using Xunit;
using SR = TypeRight.Sequences.SequenceRenumberer;

public class SequenceRenumberTests
{
    // ---- 标记解析 ----

    private static SR.MarkerHit Parse(string text) { return SR.TryParseMarker(text); }

    [Theory]
    [InlineData("1. 计算", MarkerKindArabicDot, 1)]
    [InlineData("12.已知", MarkerKindArabicDot, 12)]
    [InlineData("1．全角点", MarkerKindArabicDot, 1)]
    [InlineData("  3. 前导空白", MarkerKindArabicDot, 3)]
    [InlineData("1、顿号", MarkerKindArabicComma, 1)]
    [InlineData("（1）求值", MarkerKindParenCn, 1)]
    [InlineData("(12)半角", MarkerKindParenEn, 12)]
    [InlineData("1）中文括", MarkerKindCloseCn, 1)]
    [InlineData("1)半角括", MarkerKindCloseEn, 1)]
    [InlineData("①命题", MarkerKindCircled, 1)]
    [InlineData("㉑段落", MarkerKindCircled, 21)]
    [InlineData("㊿段落", MarkerKindCircled, 50)]
    [InlineData("一、选择题", MarkerKindCnComma, 1)]
    [InlineData("十一、填空", MarkerKindCnComma, 11)]
    [InlineData("（三十五）说明", MarkerKindCnParen, 35)]
    public void Parse_KnownPatterns(string text, int kind, int value)
    {
        var h = Parse(text);
        Assert.NotNull(h);
        Assert.Equal((SR.MarkerKind)kind, h.Kind);
        Assert.Equal(value, h.Value);
    }

    private const int MarkerKindArabicDot = 0, MarkerKindArabicComma = 1, MarkerKindParenCn = 2,
        MarkerKindParenEn = 3, MarkerKindCloseCn = 4, MarkerKindCloseEn = 5, MarkerKindCircled = 6,
        MarkerKindCnComma = 7, MarkerKindCnParen = 8;

    [Theory]
    [InlineData("1.5 米")]            // 小数
    [InlineData("3.14估算")]          // 小数紧跟
    [InlineData("2026 年")]           // 年份（无装饰）
    [InlineData("2026.11.10 日程")]   // 日期
    [InlineData("说明：如下")]         // 普通文字
    [InlineData("十年以上")]           // 中文数字无顿号
    [InlineData("二三、混排")]         // 非法中文数字
    [InlineData("（2026）年报")]       // 括号年份超 3 位
    [InlineData("")]                  // 空
    [InlineData("   ")]               // 全空白
    [InlineData("十年、混合")]         // 十年=10 合法？「十」后跟「年」截断 → 无顿号紧邻 → 不匹配
    public void Parse_Rejects(string text)
    {
        Assert.Null(Parse(text));
    }

    [Fact]
    public void Parse_KeepsOriginalDecoration()
    {
        var dot = Parse("1. 甲");
        Assert.Equal(".", dot.Suffix);
        var fullDot = Parse("1．乙");
        Assert.Equal("．", fullDot.Suffix);
        var paren = Parse("（3）丙");
        Assert.Equal("（", paren.Prefix);
        Assert.Equal("）", paren.Suffix);
        Assert.Equal(3, paren.Length);
    }

    // ---- 中文数字 ----

    [Fact]
    public void Cn_Roundtrip_1_99()
    {
        for (int n = 1; n <= 99; n++)
            Assert.Equal(n, SR.CnToInt(SR.IntToCn(n)));
    }

    [Theory]
    [InlineData("一", 1)]
    [InlineData("十", 10)]
    [InlineData("十一", 11)]
    [InlineData("二十", 20)]
    [InlineData("二十一", 21)]
    [InlineData("九十九", 99)]
    public void CnToInt_Known(string s, int v) { Assert.Equal(v, SR.CnToInt(s)); }

    [Theory]
    [InlineData("零")]
    [InlineData("两")]
    [InlineData("百")]
    [InlineData("十二三")]
    [InlineData("一百")]
    public void CnToInt_Rejects(string s) { Assert.Null(SR.CnToInt(s)); }

    [Fact]
    public void Circled_Range()
    {
        Assert.Equal('①', SR.CircledOf(1));
        Assert.Equal('⑳', SR.CircledOf(20));
        Assert.Equal('㉑', SR.CircledOf(21));
        Assert.Equal('㊿', SR.CircledOf(50));
        Assert.Equal(50, SR.CircledValue('㊿'));
        Assert.Null(SR.CircledValue('A'));
    }

    // ---- 计划（流划分 + 断序判定）----

    private static List<SR.SeqEntry> Entries(params int[] values)
    {
        var list = new List<SR.SeqEntry>();
        for (int i = 0; i < values.Length; i++)
            list.Add(new SR.SeqEntry { Ordinal = i, Hit = new SR.MarkerHit { Kind = SR.MarkerKind.ArabicDot, Value = values[i], Prefix = "", Suffix = "." } });
        return list;
    }

    private static List<SR.SeqEntry> EntriesCircled(params int[] values)
    {
        var list = new List<SR.SeqEntry>();
        for (int i = 0; i < values.Length; i++)
            list.Add(new SR.SeqEntry { Ordinal = i, Hit = new SR.MarkerHit { Kind = SR.MarkerKind.Circled, Value = values[i], Prefix = "", Suffix = "" } });
        return list;
    }

    private static Dictionary<int, string> ApplyPlan(List<SR.SeqEntry> entries, HashSet<int> scope = null)
    {
        var map = new Dictionary<int, string>();
        var plan = SR.PlanCore(entries, scope);
        foreach (var c in plan.Changes) map[c.Ordinal] = c.NewText;
        return map;
    }

    [Fact]
    public void Plan_HealthySequenceUntouched()
    {
        Assert.Empty(ApplyPlan(Entries(1, 2, 3, 4, 5)));
    }

    [Fact]
    public void Plan_GapRenumbered()
    {
        var map = ApplyPlan(Entries(1, 2, 4, 5));          // 删了 3
        Assert.Equal("3.", map[2]);
        Assert.Equal("4.", map[3]);
        Assert.Equal(2, map.Count);
    }

    [Fact]
    public void Plan_DuplicateRenumbered()
    {
        var map = ApplyPlan(Entries(1, 2, 2, 4));          // 插入重号：2,2 → 2,3，原 4 恰好归位
        Assert.Equal("3.", map[2]);
        Assert.Equal(1, map.Count);
    }

    [Fact]
    public void Plan_TranspositionRenumbered()
    {
        var map = ApplyPlan(Entries(1, 3, 2, 4));          // 乱序
        Assert.Equal("2.", map[1]);
        Assert.Equal("3.", map[2]);
    }

    [Fact]
    public void Plan_RestartAtOneSplitsRuns()              // 大题重新起号：两段各自健康
    {
        Assert.Empty(ApplyPlan(Entries(1, 2, 3, 1, 2, 3)));
    }

    [Fact]
    public void Plan_BrokenSecondSectionRenumbered()       // 第二大题断了：只修第二大题
    {
        var map = ApplyPlan(Entries(1, 2, 3, 1, 4));
        Assert.Equal("2.", map[4]);
        Assert.Equal(1, map.Count);
    }

    [Fact]
    public void Plan_SingletonUntouched()                  // 单条目不动（可能是行内序号的续文）
    {
        Assert.Empty(ApplyPlan(Entries(7)));
        Assert.Empty(ApplyPlan(Entries(1, 2, 1)));         // 尾部孤立 1 也按新流处理 → 单条目不动
    }

    [Fact]
    public void Plan_DeletedHeadRebases()                  // 删了第 1 条：2,3 → 1,2
    {
        var map = ApplyPlan(Entries(2, 3, 4));
        Assert.Equal("1.", map[0]);
        Assert.Equal("2.", map[1]);
        Assert.Equal("3.", map[2]);
    }

    [Fact]
    public void Plan_KindsIndependent()                    // 1. 与 （1） 互不干扰
    {
        var entries = new List<SR.SeqEntry>
        {
            new SR.SeqEntry { Ordinal = 0, Hit = Parse("1. 甲") },
            new SR.SeqEntry { Ordinal = 1, Hit = Parse("（1）子") },
            new SR.SeqEntry { Ordinal = 2, Hit = Parse("（2）子") },
            new SR.SeqEntry { Ordinal = 3, Hit = Parse("2. 乙") },
        };
        Assert.Empty(ApplyPlan(entries));                  // 各自连续
    }

    [Fact]
    public void Plan_CircledOverflowSkipped()              // ① 序列 >50 无字符 → 整流跳过
    {
        var values = new int[51];
        for (int i = 0; i < 51; i++) values[i] = i + 1 == 51 ? 99 : i + 1;  // 1..50 连续，第 51 项错标 99 → 断序
        var plan = SR.PlanCore(EntriesCircled(values));
        Assert.Empty(plan.Changes);
        Assert.Equal(1, plan.Skipped.Count);
    }

    [Fact]
    public void Plan_CircledNormalGapFixed()
    {
        var map = ApplyPlan(EntriesCircled(1, 2, 4));
        Assert.Equal("③", map[2]);
    }

    [Fact]
    public void Plan_ScopeFiltersWholeRunIn()
    {
        var scope = new HashSet<int> { 2 };                 // 选区只落在第 3 项
        var map = ApplyPlan(Entries(1, 2, 4, 5), scope);   // 该流含选区 → 整流处理
        Assert.Equal("3.", map[2]);
        Assert.Equal("4.", map[3]);
    }

    [Fact]
    public void Plan_ScopeExcludesUnrelatedRun()
    {
        var entries = new List<SR.SeqEntry>();
        int ord = 0;
        foreach (int v in new[] { 1, 2, 4 })                 // 1. 序列（断序，不在选区）
            entries.Add(new SR.SeqEntry { Ordinal = ord++, Hit = new SR.MarkerHit { Kind = SR.MarkerKind.ArabicDot, Value = v, Prefix = "", Suffix = "." } });
        foreach (int v in new[] { 1, 3 })                    // （1）序列（断序，在选区）
            entries.Add(new SR.SeqEntry { Ordinal = ord++, Hit = new SR.MarkerHit { Kind = SR.MarkerKind.ParenCn, Value = v, Prefix = "（", Suffix = "）" } });
        var map = ApplyPlan(entries, new HashSet<int> { 3 });
        Assert.Equal("（2）", map[4]);
        Assert.Equal(1, map.Count);                          // 1. 序列未被碰
    }

    [Fact]
    public void Plan_MixedDecorationKeptPerItem()           // 同流混用半/全角点：各保留原装饰
    {
        var entries = new List<SR.SeqEntry>
        {
            new SR.SeqEntry { Ordinal = 0, Hit = Parse("1. 甲") },
            new SR.SeqEntry { Ordinal = 1, Hit = Parse("2．乙") },
            new SR.SeqEntry { Ordinal = 2, Hit = Parse("4. 丙") },
        };
        var map = ApplyPlan(entries);
        Assert.Equal("3.", map[2]);
    }

    [Fact]
    public void Render_CnAndCircled()
    {
        var cn = Parse("一、题");
        Assert.Equal("十二、", SR.RenderMarker(cn, 12));
        var circled = Parse("①项");
        Assert.Equal("㉒", SR.RenderMarker(circled, 22));
        Assert.Null(SR.RenderMarker(circled, 51));
    }
}
