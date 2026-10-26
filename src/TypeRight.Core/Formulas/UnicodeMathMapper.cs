using System;
using System.Collections.Generic;
using System.Text;

namespace TypeRight.Formulas
{
    /// <summary>
    /// UnicodeMath 线性串映射（PRD §6.2）：
    /// Unicode 上下标字符 → _{}/^{} 语法；√ 操作数统一为 √(…)；斜杠分式交给 BuildUp 自动堆叠。
    /// </summary>
    public static class UnicodeMathMapper
    {
        private static readonly Dictionary<char, char> Superscripts = new Dictionary<char, char>
        {
            {'⁰','0'},{'¹','1'},{'²','2'},{'³','3'},{'⁴','4'},{'⁵','5'},{'⁶','6'},{'⁷','7'},{'⁸','8'},{'⁹','9'},
            {'⁺','+'},{'⁻','-'},{'⁼','='},{'⁽','('},{'⁾',')'},
            {'ⁿ','n'},{'ᵃ','a'},{'ᵇ','b'},{'ᶜ','c'},{'ᵈ','d'},{'ᵉ','e'},{'ᶠ','f'},{'ᵍ','g'},
            {'ʰ','h'},{'ⁱ','i'},{'ʲ','j'},{'ᵏ','k'},{'ˡ','l'},{'ᵐ','m'},
            {'ᵒ','o'},
            {'ᵖ','p'},{'ʳ','r'},{'ˢ','s'},{'ᵗ','t'},{'ᵘ','u'},{'ᵛ','v'},{'ʷ','w'},{'ˣ','x'},{'ʸ','y'},{'ᶻ','z'}
        };

        private static readonly Dictionary<char, char> Subscripts = new Dictionary<char, char>
        {
            {'₀','0'},{'₁','1'},{'₂','2'},{'₃','3'},{'₄','4'},{'₅','5'},{'₆','6'},{'₇','7'},{'₈','8'},{'₉','9'},
            {'₊','+'},{'₋','-'},{'₌','='},{'₍','('},{'₎',')'},
            {'ₐ','a'},{'ₑ','e'},{'ₒ','o'},{'ₓ','x'},{'ₕ','h'},{'ₖ','k'},{'ₗ','l'},{'ₘ','m'},{'ₙ','n'},{'ₚ','p'},{'ₛ','s'},{'ₜ','t'}
        };

        /// <summary>把原始片段文本转成 UnicodeMath 线性串。</summary>
        public static string Map(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var sb = new StringBuilder(text.Length + 8);
            int i = 0, n = text.Length;
            while (i < n)
            {
                char c = text[i];

                if (Superscripts.ContainsKey(c))
                {
                    string run = ReadRun(text, ref i, Superscripts);
                    sb.Append('^').Append(run);
                    continue;
                }
                if (Subscripts.ContainsKey(c))
                {
                    string run = ReadRun(text, ref i, Subscripts);
                    sb.Append('_').Append(run);
                    continue;
                }
                if (c == '√')
                {
                    i++;
                    string operand = ReadSqrtOperand(text, ref i);
                    if (operand == null) sb.Append('√');
                    else sb.Append("√(").Append(Map(operand)).Append(')');
                    continue;
                }
                sb.Append(c);
                i++;
            }
            return sb.ToString();
        }

        private static string ReadRun(string text, ref int i, Dictionary<char, char> map)
        {
            var sb = new StringBuilder();
            while (i < text.Length && map.ContainsKey(text[i]))
            {
                sb.Append(map[text[i]]);
                i++;
            }
            if (sb.Length > 1) return "{" + sb + "}";
            return sb.ToString();
        }

        /// <summary>
        /// 读取 √ 的操作数：
        /// √(x…) / √[x…] → 返回括号内内容（统一为圆括号）；
        /// √x、√2a、√Δ → 返回单个"原子"（字母串/数字串/单符号）。
        /// 返回 null 表示 √ 后无操作数（保留原样）。
        /// </summary>
        private static string ReadSqrtOperand(string text, ref int i)
        {
            if (i >= text.Length) return null;

            if (text[i] == '(' || text[i] == '[' || text[i] == '{')
            {
                char open = text[i];
                char close = open == '(' ? ')' : (open == '[' ? ']' : '}');
                int depth = 0;
                int start = i + 1;
                while (i < text.Length)
                {
                    if (text[i] == open) depth++;
                    else if (text[i] == close)
                    {
                        depth--;
                        if (depth == 0)
                        {
                            string inner = text.Substring(start, i - start);
                            i++; // 跳过闭括号
                            return inner;
                        }
                    }
                    i++;
                }
                // 括号不闭合：回退，按原子处理
                i = start;
            }

            // 原子：连续的字母 或 连续数字（含小数点）
            int s = i;
            if (char.IsLetter(text, i))
            {
                while (i < text.Length && char.IsLetter(text, i)) i++;
            }
            else if (char.IsDigit(text, i))
            {
                while (i < text.Length && (char.IsDigit(text, i) || text[i] == '.')) i++;
            }
            else
            {
                // 单个符号（Δ、π 等）
                i++;
            }
            return text.Substring(s, i - s);
        }
    }
}
