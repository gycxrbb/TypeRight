using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TypeRight.Formulas
{
    /// <summary>
    /// 数学片段切分（PRD §6.3）。角色引擎（R10 数学密度）与公式转换器共用。
    /// 规则：数学字符集内连续片段；必须含至少一个字母或数学专用符号；
    /// 纯数字/百分比/小节号不转；英文缩略词白名单排除；中文标点天然是边界。
    /// </summary>
    public static class MathSegmenter
    {
        public struct Segment
        {
            public int Start;
            public int Length;
            public string Text;

            public override string ToString() { return Start + ":" + Text; }
        }

        /// <summary>数学专用符号：几乎只出现在公式里的字符（无需字母即可判定为数学）。</summary>
        internal const string MathSpecialChars = "√Δπ±·×÷≠≈≤≥⟹⟺↔∵∴△∑∏∫°′″→←∂∞∝∈∉⊂⊃∪∩∀∃";

        /// <summary>ASCII 运算符（片段里必须至少出现一次才能判定为公式，用于排除普通英文/编号）。</summary>
        internal const string AsciiOperators = "=+-*/<>^_|";

        /// <summary>上下标 Unicode 字符（既是数学字符，也算"运算符"特征）。</summary>
        internal const string SuperscriptChars = "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ";
        internal const string SubscriptChars = "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓₔₕₖₗₘₙₚₛₜ";

        /// <summary>白名单：整体等于这些词时不转（英文缩略词等）。</summary>
        private static readonly HashSet<string> Whitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "vs", "vs.", "etc", "e.g", "e.g.", "i.e", "i.e.", "et al", "cf", "ibid",
            "Fig", "Figs", "Eq", "Eqs", "Ref", "Refs", "No", "Nos"
        };

        // 型号/编号模式：字母串-数字串，如 COVID-19、RTX-4090
        private static readonly Regex ModelNumberPattern = new Regex("^[A-Za-z]{2,}[-–]\\d+$", RegexOptions.Compiled);

        private static readonly HashSet<char> MathSpecial = new HashSet<char>(MathSpecialChars);
        private static readonly HashSet<char> SupChars = new HashSet<char>(SuperscriptChars);
        private static readonly HashSet<char> SubChars = new HashSet<char>(SubscriptChars);
        private static readonly HashSet<char> AsciiOp = new HashSet<char>(AsciiOperators);

        private static bool IsMathChar(char c)
        {
            if (c >= 'A' && c <= 'Z') return true;
            if (c >= 'a' && c <= 'z') return true;
            if (c >= '0' && c <= '9') return true;
            if (c == ' ') return true;
            if (MathSpecial.Contains(c)) return true;
            if (SupChars.Contains(c)) return true;
            if (SubChars.Contains(c)) return true;
            if (AsciiOp.Contains(c)) return true;
            // ASCII 标点：仅公式语境常见者
            switch (c)
            {
                case '(':
                case ')':
                case '[':
                case ']':
                case '{':
                case '}':
                case ',':
                case '.':
                case '%':
                case '\\':
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>返回可作为公式候选的片段（含边界与长度，文本已去首尾空格）。</summary>
        public static List<Segment> FindSegments(string text)
        {
            var result = new List<Segment>();
            if (string.IsNullOrEmpty(text)) return result;

            int i = 0, n = text.Length;
            while (i < n)
            {
                if (!IsMathChar(text[i])) { i++; continue; }
                int start = i;
                while (i < n && IsMathChar(text[i])) i++;
                string raw = text.Substring(start, i - start);
                string trimmed = raw.Trim(' ');
                if (trimmed.Length == 0) continue;
                // trim 后重新计算偏移
                int trimLeft = raw.Length - raw.TrimStart(' ').Length;
                int segStart = start + trimLeft;
                if (IsFormulaCandidate(trimmed))
                    result.Add(new Segment { Start = segStart, Length = trimmed.Length, Text = trimmed });
            }
            return result;
        }

        /// <summary>片段是否达到公式判定门槛。</summary>
        public static bool IsFormulaCandidate(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            if (Whitelist.Contains(s)) return false;
            if (ModelNumberPattern.IsMatch(s)) return false;

            bool hasLetterOrSpecial = false;
            bool hasOperatorOrSubSup = false;
            foreach (char c in s)
            {
                if (c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z') hasLetterOrSpecial = true;
                else if (MathSpecial.Contains(c) || SupChars.Contains(c) || SubChars.Contains(c)) { hasLetterOrSpecial = true; hasOperatorOrSubSup = true; }
                else if (AsciiOp.Contains(c)) hasOperatorOrSubSup = true;
            }
            // 必须含字母或数学专用符号，且含运算符/上下标，排除纯英文词与"Office 2016"类
            if (!hasLetterOrSpecial || !hasOperatorOrSubSup) return false;
            // 纯点号/逗号结尾的编号如 "6.1."（无字母）已被上面拦截；补一道：全数字带 % 也不转
            return true;
        }

        /// <summary>角色引擎用：对整段文本的数学特征分析。density = 数学字符占比（去空格）。</summary>
        public static MathAnalysis Analyze(string text)
        {
            string t = text ?? "";
            int total = 0, math = 0;
            bool hasEquals = false;
            foreach (char c in t)
            {
                if (c == ' ' || c == '\r' || c == '\n' || c == '\a') continue;
                total++;
                if (c == '=') hasEquals = true;
                if (IsMathChar(c)) math++;
            }
            var analysis = new MathAnalysis
            {
                Density = total == 0 ? 0 : (double)math / total,
                HasEquals = hasEquals
            };
            // 独立成行：去掉首尾空格后整行是一个数学候选片段
            string trimmed = t.Trim(' ', '\r', '\n', '\a');
            var segs = FindSegments(trimmed);
            analysis.Standalone = segs.Count == 1 && segs[0].Start == 0 && segs[0].Length == trimmed.Length;
            return analysis;
        }

        public struct MathAnalysis
        {
            public double Density;
            public bool HasEquals;
            public bool Standalone;
        }
    }
}
