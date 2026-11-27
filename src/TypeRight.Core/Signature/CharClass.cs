using System;

namespace TypeRight.Signature
{
    /// <summary>
    /// 一段文本的"主导字符类型"（2026-11 新增维度）。
    ///
    /// 为什么需要：用户实测（tests/TypeRight.Acceptance charclass 探针）——
    /// 文档里把数字/字母也强行设成同一字体字号后，纯数字段、纯字母段、中英混排段
    /// 与中文正文段在**现有签名下全部互相判为同类**（字体/字号/加粗/颜色全同）。
    /// 批量改格式时数字字母段被一并卷进中文正文组。
    ///
    /// 为什么不能靠 Font.NameAscii 间接判断：实测 Word 会按内容自动归一化——
    /// 纯中文段没有西文字符时 NameAscii 回落到中文字体（宋体），
    /// 有西文字符时才是 Times New Roman。这是渲染回落，不代表作者意图；
    /// 且主题字体（+中文正文）下中西文名可能完全相同。必须直接看**字符本身**。
    ///
    /// 粒度口径（用户确认）：按"主导字符类型"，不按"是否含某类字符"。
    /// 即 第1题 → 以中文为主（汉字占比最高），归 Cjk；
    ///     ABC123 → 无汉字，字母与数字各半 → Mixed；纯 ABC → Latin；纯 123 → Digit。
    /// </summary>
    public enum CharClass
    {
        /// <summary>无可见字符（空段/纯空白/纯标点）——不与任何值相等（宁可少命中）。</summary>
        None = 0,
        /// <summary>以汉字为主（含中日韩统一表意文字）。</summary>
        Cjk = 1,
        /// <summary>以拉丁字母为主（含全角字母）。</summary>
        Latin = 2,
        /// <summary>以阿拉伯数字为主。</summary>
        Digit = 3,
        /// <summary>两类及以上占比相同、无主导者（如 AB12）。</summary>
        Mixed = 4
    }

    /// <summary>主导字符类型的判定（纯逻辑，无 Word 依赖，xUnit 直测）。</summary>
    public static class CharClassDetector
    {
        /// <summary>按加权占比取主导者；并列最多 → Mixed。
        /// 标点/空白/符号不参与计数（它们不携带"这是中文还是英文"的语义）。</summary>
        public static CharClass Detect(string text)
        {
            if (string.IsNullOrEmpty(text)) return CharClass.None;

            int cjk = 0, latin = 0, digit = 0;
            foreach (char c in text)
            {
                if (IsCjk(c)) cjk++;
                else if (IsLatin(c)) latin++;
                else if (IsDigit(c)) digit++;
                // 其余（空白/标点/符号/公式标记）不计
            }

            if (cjk + latin + digit == 0) return CharClass.None;

            // 判定顺序（2026-11 依据真实文档「考察报告-商雨萤.docx」校准）：
            //
            // ① 数字压过汉字 → 数字段。
            //    实测案例：「2026年11月26日」汉字 3、数字 8。这是**日期段**，
            //    与正文（汉字 114/数字 49）在字体字号上完全相同，只能靠这条区分。
            //    若按"有汉字即中文"，日期段会被卷进正文组被一起改格式（用户实测反馈的问题）。
            // ② 其余情况有汉字 → 中文。中文夹英文词（"Python编程基础教程" 6:6）仍归中文，
            //    符合"这是中文正文"的直觉——英文词不该让整段脱离中文组。
            // ③ 无汉字：拉丁 vs 数字 按数量比较（同权，避免 "ABC123" 被单方面判成字母）。
            if (cjk > 0)
                return digit > cjk ? CharClass.Digit : CharClass.Cjk;

            if (latin > digit) return CharClass.Latin;
            if (digit > latin) return CharClass.Digit;
            return CharClass.Mixed;
        }

        /// <summary>汉字判定：基本区 U+4E00–U+9FFF、扩展 A U+3400–U+4DBF、兼容区 U+F900–U+FAFF。
        /// 不含假名/谚文（本产品面向中文文档，日韩字符按"其他"处理不参与主导判定）。</summary>
        public static bool IsCjk(char c)
        {
            return (c >= '\u4E00' && c <= '\u9FFF')
                || (c >= '\u3400' && c <= '\u4DBF')
                || (c >= '\uF900' && c <= '\uFAFF');
        }

        /// <summary>拉丁字母：ASCII a-z/A-Z + 全角 Ａ-Ｚ/ａ-ｚ。</summary>
        public static bool IsLatin(char c)
        {
            return (c >= 'A' && c <= 'Z')
                || (c >= 'a' && c <= 'z')
                || (c >= '\uFF21' && c <= '\uFF3A')
                || (c >= '\uFF41' && c <= '\uFF5A');
        }

        /// <summary>阿拉伯数字：ASCII 0-9 + 全角 ０-９。</summary>
        public static bool IsDigit(char c)
        {
            return (c >= '0' && c <= '9') || (c >= '\uFF10' && c <= '\uFF19');
        }

        /// <summary>中文名（日志/提示用）。</summary>
        public static string Name(CharClass cc)
        {
            switch (cc)
            {
                case CharClass.Cjk: return "中文";
                case CharClass.Latin: return "字母";
                case CharClass.Digit: return "数字";
                case CharClass.Mixed: return "混合";
                default: return "无";
            }
        }

        /// <summary>单个字符的类型（None = 标点/空白/符号，不参与任何一类）。</summary>
        public static CharClass ClassOf(char c)
        {
            if (IsCjk(c)) return CharClass.Cjk;
            if (IsLatin(c)) return CharClass.Latin;
            if (IsDigit(c)) return CharClass.Digit;
            return CharClass.None;
        }

        /// <summary>一个连续的同类字符片段（相对段落起点偏移）。
        /// 用于"只对某类字符刷格式/高亮"，而不是整段。</summary>
        public struct Run
        {
            public int Start;      // 相对偏移
            public int Length;
            public CharClass Class;
            public int End { get { return Start + Length; } }
        }

        /// <summary>把文本按字符类型切成连续片段（纯字符串运算，零 COM 成本）。
        /// 例："商雨萤同志，女，2005年5月出生"
        ///   → [Cjk "商雨萤同志"][Cjk "女"][Digit "2005"][Cjk "年"][Digit "5"][Cjk "月出生"]
        /// 标点/空白不产生片段（它们不携带类型语义，也不该被单独染色）。
        ///
        /// 用途（用户 2026-11 明确的口径）：同类选择命中整段后，**只对目标类型的字符**
        /// 高亮与改格式，段内的数字/字母保持原样——"汉字归汉字、数字归数字"。</summary>
        public static System.Collections.Generic.List<Run> SplitByClass(string text)
        {
            var result = new System.Collections.Generic.List<Run>();
            if (string.IsNullOrEmpty(text)) return result;

            int i = 0;
            while (i < text.Length)
            {
                CharClass cls = ClassOf(text[i]);
                if (cls == CharClass.None) { i++; continue; }
                int j = i + 1;
                while (j < text.Length && ClassOf(text[j]) == cls) j++;
                result.Add(new Run { Start = i, Length = j - i, Class = cls });
                i = j;
            }
            return result;
        }

        /// <summary>只保留指定类型的片段（可选合并被标点隔开的同类型相邻片段——
        /// 例如"商雨萤同志，女"两个中文片段中间只是逗号，合并后少刷一次格式）。</summary>
        public static System.Collections.Generic.List<Run> FilterRuns(
            System.Collections.Generic.List<Run> runs, CharClass want, bool mergeAcrossPunctuation)
        {
            var result = new System.Collections.Generic.List<Run>();
            if (runs == null || runs.Count == 0) return result;

            Run current = default(Run);
            bool has = false;

            foreach (Run r in runs)
            {
                if (r.Class != want)
                {
                    // 遇到别的类型：先把已累积的片段收进结果，再丢弃当前累积。
                    // （旧实现只把 has 置 false 而没有 flush —— 导致前面所有片段丢失）
                    if (has) { result.Add(current); has = false; }
                    continue;
                }

                if (has && mergeAcrossPunctuation && r.Start <= current.End + 1)
                {
                    // 中间最多隔 1 个字符（标点/空白）→ 合并成一段
                    int newEnd = Math.Max(current.End, r.End);
                    current.Length = newEnd - current.Start;
                    continue;
                }

                if (has) result.Add(current);
                current = r;
                has = true;
            }
            if (has) result.Add(current);
            return result;
        }
    }
}
