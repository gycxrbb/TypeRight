using System;
using System.Collections.Generic;
using System.Text;
using NetOffice.WordApi;

namespace TypeRight.Sequences
{
    /// <summary>序列重排（2026-11：插入/删除条目后 1. ① （1） 1） 一、 等序号手动补齐的痛点）。
    /// 全文扫描段首序号标记 → 按标记类型分组 → 断序序列（缺号/重号/乱序）一键重排。
    /// 设计原则：只动"断序"的序列，连续的序列绝不碰；单条目序列不动（无法推断意图）；
    /// 断 1 处开新序列（试卷分大题重新起号是常态）；重排一律以 1 为基数。</summary>
    public static class SequenceRenumberer
    {
        // ---- 标记模型 ----

        public enum MarkerKind
        {
            ArabicDot,   // 1.  1．
            ArabicComma, // 1、
            ParenCn,     // （1）
            ParenEn,     // (1)
            CloseCn,     // 1）
            CloseEn,     // 1)
            Circled,     // ①…㊿（1-50）
            CnComma,     // 一、
            CnParen      // （一）
        }

        /// <summary>一个段首序号标记的解析结果。</summary>
        public sealed class MarkerHit
        {
            public MarkerKind Kind;
            public int Value;      // 序号数值
            public int Start;      // 段内偏移（含前导空白，Range.Text 坐标系）
            public int Length;     // 标记字符数
            public string Prefix;  // 复用原装饰（（ 或 ( 或 ""）
            public string Suffix;  // 复用原装饰（. ． 、 ） ) 或 ""）
        }

        // ---- 纯逻辑层（无 Word 依赖，可直接单测）----

        /// <summary>一条扫描命中：Ordinal 是全文扫描序号（纯逻辑层的段落身份）。</summary>
        public sealed class SeqEntry
        {
            public int Ordinal;
            public MarkerHit Hit;
        }

        public sealed class PlannedChange
        {
            public int Ordinal;
            public string NewText;
        }

        /// <summary>重排计划。Changes 为要改写的标记；Summaries 是人话摘要；Skipped 是无法重排的序列说明。</summary>
        public sealed class RenumberPlanCore
        {
            public readonly List<PlannedChange> Changes = new List<PlannedChange>();
            public readonly List<string> Summaries = new List<string>();
            public readonly List<string> Skipped = new List<string>();
        }

        /// <summary>计算重排计划。
        /// 分组规则：同类型标记按文档序成流；条目值为 1 且非首条 → 开新流（大题重新起号）。
        /// 健康流：首值 1 且严格连续 → 不动。断序流（长度 ≥2）：重排为 1..n。
        /// 单条目流不动。scopeOrdinals 非空时只处理包含选区条目的流（整流处理，含流内选区外条目）。</summary>
        public static RenumberPlanCore PlanCore(IEnumerable<SeqEntry> entries, HashSet<int> scopeOrdinals = null)
        {
            var plan = new RenumberPlanCore();
            var byKind = new Dictionary<MarkerKind, List<SeqEntry>>();
            foreach (var e in entries)
            {
                List<SeqEntry> list;
                if (!byKind.TryGetValue(e.Hit.Kind, out list)) byKind[e.Hit.Kind] = list = new List<SeqEntry>();
                list.Add(e);
            }

            foreach (var kv in byKind)
            {
                var run = new List<SeqEntry>();
                Action flush = () =>
                {
                    if (run.Count == 0) return;
                    TryRenumberRun(run, scopeOrdinals, plan);
                    run = new List<SeqEntry>();
                };

                foreach (var e in kv.Value)
                {
                    if (run.Count > 0 && e.Hit.Value == 1) flush();   // 1 起号新流
                    run.Add(e);
                }
                flush();
            }
            return plan;
        }

        private static void TryRenumberRun(List<SeqEntry> run, HashSet<int> scopeOrdinals, RenumberPlanCore plan)
        {
            if (run.Count < 2) return;                                  // 单条目不动（无法推断意图）
            if (scopeOrdinals != null)
            {
                bool any = false;
                foreach (var e in run) if (scopeOrdinals.Contains(e.Ordinal)) { any = true; break; }
                if (!any) return;                                       // 选区模式：只碰选区涉及的流
            }

            bool healthy = run[0].Hit.Value == 1;
            if (healthy)
                for (int i = 1; i < run.Count && healthy; i++)
                    if (run[i].Hit.Value != run[i - 1].Hit.Value + 1) healthy = false;
            if (healthy) return;                                        // 连续序列绝不碰

            // 重排为 1..n；超出该类型字符范围的整流跳过（不做半截重排）
            for (int i = 0; i < run.Count; i++)
            {
                string rendered = RenderMarker(run[i].Hit, i + 1);
                if (rendered == null)
                {
                    plan.Skipped.Add("「" + Describe(run[0].Hit) + "」序列 " + run.Count + " 项超出可表示范围，已跳过");
                    return;
                }
                if (i + 1 != run[i].Hit.Value)
                    plan.Changes.Add(new PlannedChange { Ordinal = run[i].Ordinal, NewText = rendered });
            }
            plan.Summaries.Add(Describe(run[0].Hit) + " " + run.Count + " 项");
        }

        private static string Describe(MarkerHit h)
        {
            switch (h.Kind)
            {
                case MarkerKind.Circled: return "① 序列";
                case MarkerKind.CnComma: return "一、序列";
                case MarkerKind.CnParen: return "（一）序列";
                case MarkerKind.ParenCn: return "（1）序列";
                case MarkerKind.ParenEn: return "(1) 序列";
                case MarkerKind.CloseCn: return "1）序列";
                case MarkerKind.CloseEn: return "1) 序列";
                case MarkerKind.ArabicComma: return "1、序列";
                default: return "1. 序列";
            }
        }

        /// <summary>按原装饰渲染新序号；超出字符范围返回 null（① 仅到 ㊿=50，中文数字仅到 九十九=99）。</summary>
        public static string RenderMarker(MarkerHit template, int value)
        {
            string num;
            switch (template.Kind)
            {
                case MarkerKind.Circled:
                    if (value > 50) return null;
                    num = CircledOf(value).ToString();
                    break;
                case MarkerKind.CnComma:
                case MarkerKind.CnParen:
                    if (value > 99) return null;
                    num = IntToCn(value);
                    break;
                default:
                    num = value.ToString();
                    break;
            }
            return template.Prefix + num + template.Suffix;
        }

        // ---- 标记解析 ----

        /// <summary>解析段首序号标记（前导空白后）。小数（1.5）、年份等误匹配已排除。</summary>
        public static MarkerHit TryParseMarker(string paraText)
        {
            if (string.IsNullOrEmpty(paraText)) return null;
            int i = 0;
            while (i < paraText.Length && IsSpace(paraText[i])) i++;
            if (i >= paraText.Length) return null;
            int start = i;

            char c = paraText[i];
            // （1） / （一）
            if (c == '（')
            {
                int close = FindChar(paraText, i + 1, '）');
                if (close > i + 1 && close - i - 1 <= 4)
                {
                    string inner = paraText.Substring(i + 1, close - i - 1);
                    int? arabic = TryArabic(inner, 3);
                    if (arabic.HasValue)
                        return Hit(MarkerKind.ParenCn, arabic.Value, start, close - start + 1, "（", "）");
                    int? cn = CnToInt(inner);
                    if (cn.HasValue)
                        return Hit(MarkerKind.CnParen, cn.Value, start, close - start + 1, "（", "）");
                }
                return null;
            }
            // (1)
            if (c == '(')
            {
                int close = FindChar(paraText, i + 1, ')');
                if (close > i + 1 && close - i - 1 <= 3)
                {
                    int? arabic = TryArabic(paraText.Substring(i + 1, close - i - 1), 3);
                    if (arabic.HasValue)
                        return Hit(MarkerKind.ParenEn, arabic.Value, start, close - start + 1, "(", ")");
                }
                return null;
            }
            // ①…㊿
            int? circled = CircledValue(c);
            if (circled.HasValue)
                return Hit(MarkerKind.Circled, circled.Value, start, 1, "", "");

            // 阿拉伯数字 + 尾随装饰 / 中文数字 + 、
            int digits = CountDigits(paraText, i, 3);
            if (digits > 0)
            {
                int value = int.Parse(paraText.Substring(i, digits));
                char after = i + digits < paraText.Length ? paraText[i + digits] : '\0';
                int len = digits + 1;
                if (after == '、') return Hit(MarkerKind.ArabicComma, value, start, len, "", "、");
                if (after == '．') return Hit(MarkerKind.ArabicDot, value, start, len, "", "．");
                if (after == '.')                                     // 小数防护：1.5 不是序号
                {
                    char next = i + digits + 1 < paraText.Length ? paraText[i + digits + 1] : '\0';
                    if (!IsDigit(next)) return Hit(MarkerKind.ArabicDot, value, start, len, "", ".");
                }
                if (after == '）') return Hit(MarkerKind.CloseCn, value, start, len, "", "）");
                if (after == ')') return Hit(MarkerKind.CloseEn, value, start, len, "", ")");
                return null;
            }
            // 一、 （无括号）
            int cnLen = CountCnDigits(paraText, i, 4);
            if (cnLen > 0 && i + cnLen < paraText.Length && paraText[i + cnLen] == '、')
            {
                int? cn = CnToInt(paraText.Substring(i, cnLen));
                if (cn.HasValue)
                    return Hit(MarkerKind.CnComma, cn.Value, start, cnLen + 1, "", "、");
            }
            return null;
        }

        private static MarkerHit Hit(MarkerKind kind, int value, int start, int length, string prefix, string suffix)
        {
            return new MarkerHit { Kind = kind, Value = value, Start = start, Length = length, Prefix = prefix, Suffix = suffix };
        }

        private static bool IsSpace(char c)
        {
            return c == ' ' || c == '　' || c == '\t' || c == ' ';
        }

        private static bool IsDigit(char c) { return c >= '0' && c <= '9'; }

        private static int CountDigits(string s, int from, int max)
        {
            int n = 0;
            while (n < max && from + n < s.Length && IsDigit(s[from + n])) n++;
            return n;
        }

        private static int? TryArabic(string s, int maxDigits)
        {
            if (s.Length == 0 || s.Length > maxDigits) return null;
            for (int i = 0; i < s.Length; i++) if (!IsDigit(s[i])) return null;
            return int.Parse(s);
        }

        private static int FindChar(string s, int from, char c)
        {
            for (int i = from; i < s.Length && i < from + 8; i++)
                if (s[i] == c) return i;
            return -1;
        }

        // ---- 中文数字（一 ~ 九十九）----

        private const string CnDigits = "零一二三四五六七八九";

        /// <summary>中文数字 → 整数（一~九十九）。非标准序号写法返回 null。</summary>
        public static int? CnToInt(string s)
        {
            if (string.IsNullOrEmpty(s) || s.Length > 4) return null;
            for (int i = 0; i < s.Length; i++)
                if (s[i] != '十' && CnDigits.IndexOf(s[i]) < 1) return null;   // 只认 一~九 + 十
            if (s.Length == 1) return s[0] == '十' ? 10 : CnDigits.IndexOf(s[0]);
            int shi = s.IndexOf('十');
            if (shi < 0) return null;                                            // 多字无十（如"二三"）非法
            string tens = s.Substring(0, shi), ones = s.Substring(shi + 1);
            int t = tens.Length == 0 ? 1 : CnDigits.IndexOf(tens[0]);
            if (tens.Length > 1 || t < 1) return null;
            if (ones.Length == 0) return t * 10;
            if (ones.Length > 1) return null;
            int o = CnDigits.IndexOf(ones[0]);
            if (o < 1) return null;
            return t * 10 + o;
        }

        /// <summary>整数 → 中文数字（1~99，标准序号写法：十一、二十、二十一）。</summary>
        public static string IntToCn(int n)
        {
            if (n < 1 || n > 99) throw new ArgumentOutOfRangeException("n");
            if (n < 10) return CnDigits[n].ToString();
            if (n == 10) return "十";
            if (n < 20) return "十" + CnDigits[n % 10];
            return CnDigits[n / 10] + "十" + (n % 10 == 0 ? "" : CnDigits[n % 10].ToString());
        }

        private static int CountCnDigits(string s, int from, int max)
        {
            int n = 0;
            while (n < max && from + n < s.Length && (s[from + n] == '十' || CnDigits.IndexOf(s[from + n]) >= 1)) n++;
            return n;
        }

        // ---- 带圈数字（①-⑳ ㉑-㉟ ㊱-㊿ → 1-50）----

        public static char CircledOf(int v)
        {
            if (v < 1 || v > 50) throw new ArgumentOutOfRangeException("v");
            if (v <= 20) return (char)(0x2460 + v - 1);
            if (v <= 35) return (char)(0x3251 + v - 21);
            return (char)(0x32B1 + v - 36);
        }

        public static int? CircledValue(char c)
        {
            if (c >= 0x2460 && c <= 0x2473) return c - 0x2460 + 1;
            if (c >= 0x3251 && c <= 0x325F) return c - 0x3251 + 21;
            if (c >= 0x32B1 && c <= 0x32BF) return c - 0x32B1 + 36;
            return null;
        }

        // ---- Word 层 ----

        public sealed class ChangeRow
        {
            public Paragraph Para;
            public int StartInPara;
            public int Length;
            public string NewText;
        }

        public sealed class RenumberPlan
        {
            public readonly List<ChangeRow> Changes = new List<ChangeRow>();
            public readonly List<string> Summaries = new List<string>();
            public readonly List<string> Skipped = new List<string>();
        }

        /// <summary>扫描全文计算重排计划。scope 非 null 时只处理与选区有序号交集的序列（整序列处理）。</summary>
        public static RenumberPlan Plan(Document doc, Range scope)
        {
            var plan = new RenumberPlan();
            var entries = new List<SeqEntry>();
            var paras = new List<Paragraph>();
            var scopeOrdinals = new HashSet<int>();
            bool scoped = scope != null;
            int scopeStart = 0, scopeEnd = 0;
            if (scoped) { scopeStart = scope.Start; scopeEnd = scope.End; }

            int ordinal = 0;
            foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
            {
                int pStart, pEnd;
                string text;
                try
                {
                    Range r = p.Range;
                    pStart = r.Start; pEnd = r.End;
                    text = r.Text ?? "";
                }
                catch { continue; }
                ordinal++;
                var hit = TryParseMarker(text);
                if (hit == null) continue;
                entries.Add(new SeqEntry { Ordinal = ordinal, Hit = hit });
                paras.Add(p);
                if (scoped && pStart < scopeEnd && pEnd > scopeStart)
                    scopeOrdinals.Add(ordinal);
            }

            var core = PlanCore(entries, scoped ? scopeOrdinals : null);
            foreach (var c in core.Changes)
            {
                int idx = entries.FindIndex(e => e.Ordinal == c.Ordinal);
                var e = entries[idx];
                plan.Changes.Add(new ChangeRow { Para = paras[idx], StartInPara = e.Hit.Start, Length = e.Hit.Length, NewText = c.NewText });
            }
            plan.Summaries.AddRange(core.Summaries);
            plan.Skipped.AddRange(core.Skipped);
            return plan;
        }

        /// <summary>应用计划：逐条改写标记文本（Paragraph.Range 每次现取，自动适应前面的长度变化）。</summary>
        public static void Apply(Document doc, RenumberPlan plan)
        {
            foreach (var c in plan.Changes)
            {
                Range pr = c.Para.Range;
                Range sub = doc.Range(pr.Start + c.StartInPara, pr.Start + c.StartInPara + c.Length);
                sub.Text = c.NewText;
            }
        }
    }
}
