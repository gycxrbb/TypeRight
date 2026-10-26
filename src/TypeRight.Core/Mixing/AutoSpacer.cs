using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace TypeRight.Mixing
{
    public struct TextEdit
    {
        /// <summary>替换区间起点（相对段落文本的字符偏移）。Length=0 表示纯插入。</summary>
        public int Position;
        public int Length;
        /// <summary>替换/插入的内容。</summary>
        public string Insert;

        public static TextEdit InsertAt(int position, string text)
        {
            return new TextEdit { Position = position, Length = 0, Insert = text };
        }

        public static TextEdit Replace(int position, int length, string text)
        {
            return new TextEdit { Position = position, Length = length, Insert = text };
        }
    }

    /// <summary>
    /// 中西文混排规则（PRD §3 mixingRules）：autoSpaceCjkLatin 与 unifyPunctuation。
    /// 规则集移植自 huacnlee/autocorrect 与 vinta/pangu.js 的核心子集（空格插入 + 标点映射的正则集）。
    /// 幂等：已处理的文本再次处理不产生新编辑。
    /// </summary>
    public static class AutoSpacer
    {
        private const string HanRange = "㐀-䶿一-鿿豈-﫿";

        // 汉字与半角字母数字之间补空格（互相两个方向）
        private static readonly Regex SpaceAfterHan = new Regex("(?<=[" + HanRange + "])(?=[A-Za-z0-9])", RegexOptions.Compiled);
        private static readonly Regex SpaceBeforeHan = new Regex("(?<=[A-Za-z0-9])(?=[" + HanRange + "])", RegexOptions.Compiled);

        // 全角标点（含中文顿号等）
        private const string CjkPunct = "，。：；！？、”“’‘（）《》【】—…·";

        private static readonly Dictionary<char, char> HalfToFull = new Dictionary<char, char>
        {
            {',', '，'}, {'.', '。'}, {':', '：'}, {';', '；'}, {'!', '！'}, {'?', '？'}
        };

        private static readonly Dictionary<char, char> FullToHalf = new Dictionary<char, char>
        {
            {'，', ','}, {'。', '.'}, {'：', ':'}, {'；', ';'}, {'！', '!'}, {'？', '?'}
        };

        public static bool IsHan(char c)
        {
            return (c >= '㐀' && c <= '䶿') || (c >= '一' && c <= '鿿') || (c >= '豈' && c <= '﫿');
        }

        private static bool IsCjkContext(char c)
        {
            return IsHan(c) || CjkPunct.IndexOf(c) >= 0;
        }

        private static bool IsAsciiAlnum(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9');
        }

        /// <summary>计算全部文本编辑（插入空格 + 标点统一），返回按位置升序。</summary>
        public static List<TextEdit> ComputeEdits(string text, bool autoSpace, Profiles.PunctuationMode punctuation)
        {
            var edits = new List<TextEdit>();
            if (string.IsNullOrEmpty(text)) return edits;

            if (autoSpace)
            {
                foreach (Match m in SpaceAfterHan.Matches(text))
                    edits.Add(TextEdit.InsertAt(m.Index, " "));
                foreach (Match m in SpaceBeforeHan.Matches(text))
                    edits.Add(TextEdit.InsertAt(m.Index, " "));
            }

            if (punctuation == Profiles.PunctuationMode.Fullwidth)
            {
                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];
                    char full;
                    if (!HalfToFull.TryGetValue(c, out full)) continue;
                    char prev = i > 0 ? text[i - 1] : '\0';
                    char next = i < text.Length - 1 ? text[i + 1] : '\0';
                    if (prev == '\0' || next == '\0') continue;
                    // 小数点/千分位保护：两侧都是数字
                    if (char.IsDigit(prev) && char.IsDigit(next)) continue;
                    // 仅在中文语境（两侧为汉字或全角标点）时全角化
                    if (IsCjkContext(prev) && IsCjkContext(next))
                        edits.Add(TextEdit.Replace(i, 1, full.ToString()));
                }
            }
            else if (punctuation == Profiles.PunctuationMode.Halfwidth)
            {
                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];
                    char half;
                    if (!FullToHalf.TryGetValue(c, out half)) continue;
                    char prev = i > 0 ? text[i - 1] : '\0';
                    char next = i < text.Length - 1 ? text[i + 1] : '\0';
                    // 纯西文语境（两侧为 ASCII 字母数字或空格）时半角化
                    bool asciiCtx(char x) { return x == '\0' || x == ' ' || IsAsciiAlnum(x); }
                    if (asciiCtx(prev) && asciiCtx(next))
                        edits.Add(TextEdit.Replace(i, 1, half.ToString()));
                }
            }

            edits.Sort((a, b) => a.Position != b.Position ? a.Position.CompareTo(b.Position) : b.Length.CompareTo(a.Length));
            return edits;
        }

        /// <summary>纯文本应用（单测与预览用）；Word 端由 FormatApplier 走 Range 写入。</summary>
        public static string Apply(string text, bool autoSpace, Profiles.PunctuationMode punctuation)
        {
            var edits = ComputeEdits(text, autoSpace, punctuation);
            if (edits.Count == 0) return text;
            var sb = new System.Text.StringBuilder(text.Length + edits.Count);
            int pos = 0;
            foreach (var e in edits)
            {
                if (e.Position < pos) continue; // 防御：跳过重叠
                sb.Append(text, pos, e.Position - pos);
                sb.Append(e.Insert);
                pos = e.Position + e.Length;
            }
            sb.Append(text, pos, text.Length - pos);
            return sb.ToString();
        }
    }
}
