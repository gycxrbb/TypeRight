using System;
using System.Collections.Generic;

namespace TypeRight.Formulas
{
    /// <summary>
    /// 结构化公式构建的纯逻辑部分：识别顶层分式与根式 span（供 FormulaConverter 经
    /// OMathFunctions.Add(wdOMathFunctionFrac/Rad) 落地——BuildUp 只构建上下标，不建分式/根式，探针实测）。
    /// </summary>
    public static class StructuredMath
    {
        /// <summary>是否需要结构化构建（顶层斜杠或 √( 出现）。</summary>
        public static bool NeedsStructure(string linear)
        {
            return TopLevelSlashIndex(linear) >= 0 || linear != null && linear.Contains("√(");
        }

        /// <summary>
        /// 规范化分式操作数：形如 (A)/(B) 的顶层分式，剥去操作数外层括号
        /// （分式线本身承担分组；Word 的 Frac 函数按 "num/den" 文本切分）。
        /// </summary>
        public static string NormalizeFractionParens(string linear)
        {
            if (string.IsNullOrEmpty(linear)) return linear;
            int slash = TopLevelSlashIndex(linear);
            if (slash < 0) return linear;

            int leftStart = OperandStart(linear, slash);
            int rightEnd = OperandEnd(linear, slash);
            string left = linear.Substring(leftStart, slash - leftStart).Trim();
            string right = linear.Substring(slash + 1, rightEnd - slash - 1).Trim();

            left = StripOuterParens(left);
            right = StripOuterParens(right);

            return linear.Substring(0, leftStart) + left + "/" + right + linear.Substring(rightEnd);
        }

        /// <summary>顶层（深度 0）斜杠位置；无则 -1。</summary>
        public static int TopLevelSlashIndex(string linear)
        {
            if (linear == null) return -1;
            int depth = 0;
            for (int i = 0; i < linear.Length; i++)
            {
                char c = linear[i];
                if (c == '(' || c == '[' || c == '{') depth++;
                else if (c == ')' || c == ']' || c == '}') depth--;
                else if (c == '/' && depth == 0) return i;
            }
            return -1;
        }

        /// <summary>√(...) 的 span 列表（含 √ 与闭合括号），按出现顺序。</summary>
        public static List<KeyValuePair<int, int>> RadicalSpans(string linear)
        {
            var spans = new List<KeyValuePair<int, int>>();
            if (linear == null) return spans;
            for (int i = 0; i < linear.Length - 1; i++)
            {
                if (linear[i] != '√') continue;
                if (linear[i + 1] != '(') continue;
                int depth = 0;
                for (int j = i + 1; j < linear.Length; j++)
                {
                    char c = linear[j];
                    if (c == '(') depth++;
                    else if (c == ')')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            spans.Add(new KeyValuePair<int, int>(i, j + 1));
                            i = j;
                            break;
                        }
                    }
                }
            }
            return spans;
        }

        /// <summary>顶层 “num/den” 操作数 span（含斜杠）；无顶层斜杠返回 null。</summary>
        public static KeyValuePair<int, int>? FractionSpan(string linear)
        {
            int slash = TopLevelSlashIndex(linear);
            if (slash < 0) return null;
            return new KeyValuePair<int, int>(OperandStart(linear, slash), OperandEnd(linear, slash));
        }

        /// <summary>分式左操作数起点：向前扫到最近的顶层运算符或串首。</summary>
        private static int OperandStart(string s, int slash)
        {
            int depth = 0;
            for (int i = slash - 1; i >= 0; i--)
            {
                char c = s[i];
                if (c == ')' || c == ']' || c == '}') depth++;
                else if (c == '(' || c == '[' || c == '{') depth--;
                else if (depth == 0 && IsSeparator(c)) return i + 1;
            }
            return 0;
        }

        /// <summary>分式右操作数终点：向后扫到最近的顶层运算符或串尾。</summary>
        private static int OperandEnd(string s, int slash)
        {
            int depth = 0;
            for (int i = slash + 1; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '(' || c == '[' || c == '{') depth++;
                else if (c == ')' || c == ']' || c == '}') depth--;
                else if (depth == 0 && IsSeparator(c)) return i;
            }
            return s.Length;
        }

        private static bool IsSeparator(char c)
        {
            switch (c)
            {
                case '+': case '-': case '±': case '=': case '<': case '>':
                case '≤': case '≥': case '≠': case ',': case ' ':
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>剥去整串的最外层配对括号（仅当确实包裹整串时）。</summary>
        public static string StripOuterParens(string s)
        {
            while (s.Length >= 2 && s[0] == '(' && s[s.Length - 1] == ')')
            {
                int depth = 0;
                bool outer = true;
                for (int i = 0; i < s.Length; i++)
                {
                    char c = s[i];
                    if (c == '(') depth++;
                    else if (c == ')')
                    {
                        depth--;
                        if (depth == 0 && i < s.Length - 1) { outer = false; break; }
                    }
                }
                if (!outer) break;
                s = s.Substring(1, s.Length - 2).Trim();
            }
            return s;
        }
    }
}
