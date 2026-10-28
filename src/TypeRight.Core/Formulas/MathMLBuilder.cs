using System;
using System.Collections.Generic;
using System.Text;

namespace TypeRight.Formulas
{
    /// <summary>
    /// 线性 UnicodeMath（UnicodeMathMapper 输出：x_1^2、√(…)、a/b、± Δ ⟺ …）→ MathML。
    /// TC-19 管线（PRD §6.2.1）：MathML → 剪贴板 UnicodeText → Word 粘贴自动转 OMML（探针 P4 实证）。
    /// 语义：每个顶层 '/' → mfrac（操作数为相邻因子，括号组整体作因子并剥外壳）；
    /// √(E) → msqrt；base_sub^sup → msub/msup/msubsup；其余字符按 mi/mn/mo 平铺。
    /// </summary>
    public static class MathMLBuilder
    {
        private const string MathNs = "http://www.w3.org/1998/Math/MathML";

        public static string Build(string linear)
        {
            if (string.IsNullOrEmpty(linear)) throw new ArgumentException("空公式");
            var p = new Parser(linear.Replace("−", "-"));
            Node root = p.ParseRow(null);
            var sb = new StringBuilder();
            sb.Append("<math xmlns=\"").Append(MathNs).Append("\">");
            root.Write(sb);
            sb.Append("</math>");
            return sb.ToString();
        }

        // ---- 节点模型 ----

        private abstract class Node
        {
            public abstract void Write(StringBuilder sb);

            protected static string Esc(string s)
            {
                return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
            }
        }

        private sealed class Row : Node
        {
            public readonly List<Node> Items = new List<Node>();
            public override void Write(StringBuilder sb)
            {
                if (Items.Count == 1) { Items[0].Write(sb); return; }
                sb.Append("<mrow>");
                foreach (var n in Items) n.Write(sb);
                sb.Append("</mrow>");
            }
        }

        private sealed class Frac : Node
        {
            private readonly Node _num, _den;
            public Frac(Node num, Node den) { _num = num; _den = den; }
            public override void Write(StringBuilder sb)
            {
                sb.Append("<mfrac>");
                sb.Append("<mrow>"); _num.Write(sb); sb.Append("</mrow>");
                sb.Append("<mrow>"); _den.Write(sb); sb.Append("</mrow>");
                sb.Append("</mfrac>");
            }
        }

        private sealed class Sqrt : Node
        {
            private readonly Node _inner;
            public Sqrt(Node inner) { _inner = inner; }
            public override void Write(StringBuilder sb)
            {
                sb.Append("<msqrt>");
                sb.Append("<mrow>"); _inner.Write(sb); sb.Append("</mrow>");
                sb.Append("</msqrt>");
            }
        }

        private sealed class Script : Node
        {
            private readonly Node _base, _sub, _sup;
            public Script(Node b, Node sub, Node sup) { _base = b; _sub = sub; _sup = sup; }
            public override void Write(StringBuilder sb)
            {
                if (_sub != null && _sup != null)
                {
                    sb.Append("<msubsup>");
                    _base.Write(sb); sb.Append("<mrow>"); _sub.Write(sb); sb.Append("</mrow><mrow>"); _sup.Write(sb); sb.Append("</mrow>");
                    sb.Append("</msubsup>");
                }
                else if (_sub != null)
                {
                    sb.Append("<msub>");
                    _base.Write(sb); sb.Append("<mrow>"); _sub.Write(sb); sb.Append("</mrow>");
                    sb.Append("</msub>");
                }
                else
                {
                    sb.Append("<msup>");
                    _base.Write(sb); sb.Append("<mrow>"); _sup.Write(sb); sb.Append("</mrow>");
                    sb.Append("</msup>");
                }
            }
        }

        private sealed class Tok : Node
        {
            public readonly string Text;
            public readonly string Kind;   // mi / mn / mo
            public bool StretchyBar;
            public Tok(string text, string kind) { Text = text; Kind = kind; }
            public override void Write(StringBuilder sb)
            {
                sb.Append('<').Append(Kind);
                if (StretchyBar) sb.Append(" stretchy=\"false\"");
                sb.Append('>').Append(Esc(Text)).Append("</").Append(Kind).Append('>');
            }
        }

        // ---- 解析器 ----

        private sealed class Parser
        {
            private readonly string _s;
            private int _i;

            public Parser(string s) { _s = s; }

            private char Cur { get { return _i < _s.Length ? _s[_i] : '\0'; } }
            private bool Eat(char c) { if (Cur == c) { _i++; return true; } return false; }

            /// <summary>解析到串尾或闭括号（close='\0' 表示到尾）。</summary>
            public Row ParseRow(char? close)
            {
                var row = new Row();
                while (_i < _s.Length)
                {
                    char c = Cur;
                    if (close.HasValue && c == close.Value) break;
                    if (c == ' ') { _i++; continue; }

                    if (IsOperator(c))
                    {
                        row.Items.Add(Operator(c));
                        _i++;
                        continue;
                    }

                    Node f = ParseFactor();
                    if (f == null) { _i++; continue; } // 未知字符跳过

                    // 分式判定需容忍斜杠两侧空格（√Δ / |a|）
                    int save = _i;
                    while (Cur == ' ') _i++;
                    if (Cur == '/')
                    {
                        _i++;
                        while (Cur == ' ') _i++;
                        Node g = ParseFactor();
                        if (g == null) g = new Tok("?", "mi");
                        row.Items.Add(new Frac(f, g));
                    }
                    else
                    {
                        _i = save;   // 回退吃掉的空格（作为普通项间空白）
                        row.Items.Add(f);
                    }
                }
                return row;
            }

            /// <summary>因子 + 上下标后缀。</summary>
            private Node ParseFactor()
            {
                Node b = ParseAtom();
                if (b == null) return null;

                Node sub = null, sup = null;
                while (Cur == '_' || Cur == '^')
                {
                    bool isSub = Cur == '_';
                    _i++;
                    Node arg = ParseScriptArg();
                    if (isSub) sub = arg; else sup = arg;
                }
                if (sub != null || sup != null) return new Script(b, sub, sup);
                return b;
            }

            private Node ParseScriptArg()
            {
                if (Eat('{'))
                {
                    var r = ParseRow('}');
                    Eat('}');
                    return r;
                }
                // 单字符参数
                if (Cur == '(')
                {
                    _i++;
                    var r = ParseRow(')');
                    Eat(')');
                    return r;
                }
                return ParseAtom();
            }

            private Node ParseAtom()
            {
                char c = Cur;
                if (c == '\0') return null;

                if (c == '|')
                {
                    // 绝对值/行列式定界组：|E| → mo(|) E mo(|)（stretchy=false）
                    _i++;
                    var inner = ParseRow('|');
                    Eat('|');
                    var g = new Row();
                    g.Items.Add(Fence('|'));
                    g.Items.Add(inner);
                    g.Items.Add(Fence('|'));
                    return g;
                }

                if (c == '(' || c == '[' || c == '{')
                {
                    char open = c;
                    char close = open == '(' ? ')' : (open == '[' ? ']' : '}');
                    _i++;
                    var inner = ParseRow(close);
                    if (!Eat(close)) { /* 不闭合：按已解析内容处理 */ }
                    var g = new Row();
                    g.Items.Add(Fence(open));
                    g.Items.Add(inner);
                    g.Items.Add(Fence(close));
                    return g;
                }

                if (c == '√')
                {
                    _i++;
                    if (Cur == '(' || Cur == '[')
                    {
                        char open = Cur; char close = open == '(' ? ')' : ']';
                        _i++;
                        var inner = ParseRow(close);
                        Eat(close);
                        return new Sqrt(inner);
                    }
                    Node atom = ParseAtom();
                    if (atom == null) atom = new Tok("?", "mi");
                    return new Sqrt(atom);
                }

                if (char.IsLetter(c))
                {
                    var sb = new StringBuilder();
                    while (_i < _s.Length && char.IsLetter(_s[_i]) && !IsOperator(_s[_i]))
                    {
                        sb.Append(_s[_i]); _i++;
                        // 连续字母按单变量序列：保留多字母为一个 mi（Word 会按 italic 变量渲染）
                    }
                    return new Tok(sb.ToString(), "mi");
                }

                if (char.IsDigit(c))
                {
                    var sb = new StringBuilder();
                    while (_i < _s.Length && (char.IsDigit(_s[_i]) || _s[_i] == '.'))
                    {
                        sb.Append(_s[_i]); _i++;
                    }
                    return new Tok(sb.ToString(), "mn");
                }

                return null;
            }

            private static Tok Operator(char c)
            {
                if (c == '|')
                {
                    return new Tok("|", "mo") { StretchyBar = true };
                }
                return new Tok(c.ToString(), "mo");
            }

            private static Tok Fence(char c)
            {
                if (c == '|') return new Tok("|", "mo") { StretchyBar = true };
                return new Tok(c.ToString(), "mo");
            }

            private static bool IsOperator(char c)
            {
                switch (c)
                {
                    case '+': case '-': case '±': case '×': case '÷': case '=':
                    case '≠': case '≈': case '≤': case '≥': case '<': case '>':
                    case '·': case '∝': case '∈': case '∉': case '∋':
                    case '⟺': case '⇔': case '⟹': case '⇒': case '↔': case '→': case '←':
                    case '∵': case '∴': case ',': case '∥': case '⊥':
                        return true;
                    default:
                        return false;
                }
            }
        }
    }
}
