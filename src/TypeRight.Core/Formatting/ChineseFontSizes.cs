using System;
using System.Collections.Generic;

namespace TypeRight.Formatting
{
    /// <summary>
    /// 中文字号 ↔ 磅值换算（Word「开始」选项卡字号框的中文命名）。
    /// 用户诉求：浮条字号只显示「14」这样的数字，与 Word 原生「四号」割裂——
    /// 统一在中文字号标准档位上显示中文名，非标准档位（如 12.5）仍显示数值。
    ///
    /// 纯逻辑层：不依赖 Word，xUnit 可直接测（与 SequenceRenumberer 同一取舍）。
    /// 磅值口径与 Word 中文版完全一致：
    /// 初号42 小初36 一号26 小一24 二号22 小二18 三号16 小三15
    /// 四号14 小四12 五号10.5 小五9 六号7.5 小六6.5 七号5.5 八号5
    /// </summary>
    public static class ChineseFontSizes
    {
        /// <summary>一个中文字号档位。</summary>
        public struct Entry
        {
            public readonly string Name;
            public readonly double Pt;

            public Entry(string name, double pt) { Name = name; Pt = pt; }
        }

        /// <summary>中文字号表（由大到小，与 Word 下拉顺序一致）。</summary>
        public static readonly Entry[] Table =
        {
            new Entry("初号", 42),  new Entry("小初", 36),
            new Entry("一号", 26),  new Entry("小一", 24),
            new Entry("二号", 22),  new Entry("小二", 18),
            new Entry("三号", 16),  new Entry("小三", 15),
            new Entry("四号", 14),  new Entry("小四", 12),
            new Entry("五号", 10.5), new Entry("小五", 9),
            new Entry("六号", 7.5), new Entry("小六", 6.5),
            new Entry("七号", 5.5), new Entry("八号", 5)
        };

        /// <summary>磅值比较容差（Word 字号是浮点，10.5 这类值读取可能有末位抖动）。</summary>
        private const double Epsilon = 0.01;

        /// <summary>磅值 → 中文字号名；非标准档位返回 null。</summary>
        public static string NameOf(double pt)
        {
            if (double.IsNaN(pt) || double.IsInfinity(pt)) return null;
            for (int i = 0; i < Table.Length; i++)
                if (Math.Abs(Table[i].Pt - pt) < Epsilon) return Table[i].Name;
            return null;
        }

        /// <summary>中文字号名 → 磅值；不认识的名字返回 null。</summary>
        public static double? PtOf(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            string key = name.Trim();
            for (int i = 0; i < Table.Length; i++)
                if (string.Equals(Table[i].Name, key, StringComparison.Ordinal)) return Table[i].Pt;
            return null;
        }

        /// <summary>磅值文本（10.5 保留一位，14 不带小数点）。</summary>
        public static string PtText(double pt)
        {
            if (double.IsNaN(pt) || double.IsInfinity(pt)) return "—";
            return pt.ToString("0.#");
        }

        /// <summary>显示用文本：中文字号档位显示「四号」，其余显示数值（12.5）。
        /// 这是字号框唯一的显示口径——按钮、下拉、提示三处必须一致。</summary>
        public static string Label(double pt)
        {
            return NameOf(pt) ?? PtText(pt);
        }

        /// <summary>完整说明（悬停提示用）：「四号（14 磅）」/「12.5 磅」。</summary>
        public static string Description(double pt)
        {
            string name = NameOf(pt);
            return name != null
                ? name + "（" + PtText(pt) + " 磅）"
                : PtText(pt) + " 磅";
        }

        /// <summary>Word 字号下拉里的候选档位（中文字号优先，随后数值档），
        /// 供字号弹层展示双轨对照。</summary>
        public static IList<Entry> NamedEntries()
        {
            return new List<Entry>(Table);
        }

        // ─────────────── 手输解析（浮条字号框可打字录入）───────────────

        /// <summary>Word 允许的字号范围（磅）。超出视为输入错误而不是截断——
        /// 用户打 "0" 或 "2000" 时按最小值/最大值悄悄改掉，比直接拒绝更让人困惑。</summary>
        public const double MinPt = 1;
        public const double MaxPt = 1638;

        /// <summary>解析用户手输的字号文本。三类合法写法：
        /// ① 中文字号名：「四号」「小四」；② 纯数字磅值：「14」「12.5」「14 磅」「14pt」；
        /// ③ 全角数字与中文字号名混排（中文输入法下常见）。
        /// 解析失败返回 null，由调用方决定提示方式；绝不静默套用某个默认值。</summary>
        public static double? TryParse(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;

            // 全角数字/小数点/空白 → 半角（中文输入法下用户敲的是全角）
            string text = NormalizeNumericWidth(input).Trim();
            if (text.Length == 0) return null;

            // ① 中文字号名：先精确匹配，再容忍末尾多余字符（「四号磅」这类手输残留）
            //    注意所有字号名都是两字，所以不存在"前缀残缺"的情形——不做部分匹配，
            //    否则「四」这种半截输入会被猜成某个档位，比直接判非法更糟。
            double? byName = PtOf(text);
            if (byName.HasValue) return byName;

            foreach (Entry e in Table)
            {
                if (text.Length > e.Name.Length && text.StartsWith(e.Name, StringComparison.Ordinal))
                    return e.Pt;
            }

            // ② 数字磅值：剥掉常见单位后缀与空白
            string numeric = StripPtSuffix(text);
            double pt;
            if (double.TryParse(numeric, System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.InvariantCulture, out pt))
            {
                if (double.IsNaN(pt) || double.IsInfinity(pt)) return null;
                if (pt < MinPt || pt > MaxPt) return null;
                return pt;
            }

            return null;
        }

        /// <summary>全角数字/小数点/加号 → 半角；全角空格转半角。其余原样保留（字号名是汉字）。</summary>
        private static string NormalizeNumericWidth(string s)
        {
            var sb = new System.Text.StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (c >= '０' && c <= '９') sb.Append((char)(c - '０' + '0'));
                else if (c == '．') sb.Append('.');
                else if (c == '＋') sb.Append('+');
                else if (c == '\u3000') sb.Append(' ');
                else sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>剥掉磅值单位后缀（「磅」「pt」「PT」及紧邻空白）。</summary>
        private static string StripPtSuffix(string s)
        {
            s = s.Trim();
            foreach (string suffix in new[] { "磅", "point", "points", "pt", "PT", "Pt", "p" })
            {
                if (s.Length > suffix.Length && s.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    string head = s.Substring(0, s.Length - suffix.Length).Trim();
                    if (head.Length > 0) { s = head; break; }
                }
            }
            return s.Trim();
        }
    }
}
