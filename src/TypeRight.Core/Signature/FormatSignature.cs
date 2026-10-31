using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formatting;

namespace TypeRight.Signature
{
    /// <summary>
    /// 格式签名（PRD v2 §2，双锚点模型）：段落**首个可见字符**与**末个可见字符**两个锚点的
    /// (中文字体, 西文字体, 字号, 加粗, 颜色) 元组 + 段落级斜体/对齐。
    /// 两个段落判同类 ⇔ 两个锚点分别相等：
    /// - 均匀段（首尾同格式）退化为旧版单锚点精确匹配；
    /// - 同构混合段（如 9 个"行首★加粗着色 + 段尾普通"的导读框）互相命中；
    /// - 混合段与首字符同格式的均匀段不命中（跨判 bug 修复）。
    /// 未知值（NaN 字号 / 杂值加粗 / null 字体）不与任何值相等——宁可少命中，不可错命中。
    /// </summary>
    public struct FormatSignature : IEquatable<FormatSignature>
    {
        // ---- 首锚点（保留原字段名：方案规则/录制等既有消费方按首锚点工作）----
        public string EastAsiaFont;   // null = 未知
        public string LatinFont;
        public double SizePt;         // NaN = 未知
        public int BoldRaw;           // -1 真 / 0 假 / 其他(9999999/杂值) = 未知
        public int ItalicRaw;
        public int ColorRgb;          // -1 = 自动；0xRRGGBB
        public string Alignment;      // "l"/"c"/"r"/"j"/"?"

        // ---- 尾锚点（段内格式不一致的段落靠它与同构混合段相认）----
        public string LastEastAsiaFont;
        public string LastLatinFont;
        public double LastSizePt;
        public int LastBoldRaw;
        public int LastColorRgb;

        public bool Bold { get { return BoldRaw == -1; } }
        public bool BoldMixed { get { return BoldRaw == WordFormatUtils.UndefinedInt; } }
        public bool Italic { get { return ItalicRaw == -1; } }

        /// <summary>首尾锚点是否一致（均匀段）。</summary>
        public bool IsUniform
        {
            get
            {
                return string.Equals(EastAsiaFont, LastEastAsiaFont, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(LatinFont, LastLatinFont, StringComparison.OrdinalIgnoreCase)
                    && SizePtEquals(SizePt, LastSizePt)
                    && BoldRaw == LastBoldRaw
                    && ColorRgb == LastColorRgb;
            }
        }

        public bool IsEmpty
        {
            get { return EastAsiaFont == null && LatinFont == null && double.IsNaN(SizePt); }
        }

        public override string ToString()
        {
            string head = string.Format("{0}/{1} {2}pt B={3} #{4} I={5} {6}",
                EastAsiaFont ?? "?", LatinFont ?? "?",
                double.IsNaN(SizePt) ? "?" : SizePt.ToString("0.#"),
                BoldRaw == -1 ? 1 : (BoldRaw == 0 ? 0 : -1),
                ColorRgb < 0 ? "自动" : ColorRgb.ToString("X6"),
                ItalicRaw,
                Alignment);
            if (IsUniform) return head;
            return head + string.Format(" ⇄尾 {0}/{1} {2}pt B={3} #{4}",
                LastEastAsiaFont ?? "?", LastLatinFont ?? "?",
                double.IsNaN(LastSizePt) ? "?" : LastSizePt.ToString("0.#"),
                LastBoldRaw == -1 ? 1 : (LastBoldRaw == 0 ? 0 : -1),
                LastColorRgb < 0 ? "自动" : LastColorRgb.ToString("X6"));
        }

        public bool Equals(FormatSignature other)
        {
            return Matches(other, IgnoreDims.None);
        }

        // Word 的 Font.Bold/Italic 会返回非标准杂值（实测混入颜色字节，如 1717853）。
        // 归一化：只有 -1 视为真，0 视为假，其余（9999999/杂值）= 未知。
        public bool BoldEffective { get { return BoldRaw == -1; } }
        public bool BoldKnown { get { return BoldRaw == -1 || BoldRaw == 0; } }
        public bool ItalicEffective { get { return ItalicRaw == -1; } }

        // 未知字号不等于任何字号（含未知==未知）：读取失败的段落不得互相"同类"（跨判 bug）。
        private static bool SizePtEquals(double a, double b)
        {
            if (double.IsNaN(a) || double.IsNaN(b)) return false;
            return Math.Abs(a - b) < 0.01;
        }

        private static bool BoldEquals(int a, int b)
        {
            bool aKnown = a == -1 || a == 0, bKnown = b == -1 || b == 0;
            if (!aKnown || !bKnown) return false;
            return a == b;
        }

        private static bool ColorEquals(int a, int b)
        {
            return a == b;   // -1(自动)==-1(自动) 相等；-1 vs 显式色 不等
        }

        private static bool FontEquals(string a, string b)
        {
            if (a == null || b == null) return false;   // 未知字体不与任何字体相等（含未知==未知）
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            return obj is FormatSignature && Equals((FormatSignature)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int h = 17;
                h = h * 31 + (EastAsiaFont == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(EastAsiaFont));
                h = h * 31 + (LatinFont == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(LatinFont));
                h = h * 31 + (double.IsNaN(SizePt) ? 0 : SizePt.GetHashCode());
                h = h * 31 + BoldRaw;
                h = h * 31 + ColorRgb;
                h = h * 31 + (double.IsNaN(LastSizePt) ? 0 : LastSizePt.GetHashCode());
                h = h * 31 + LastBoldRaw;
                h = h * 31 + LastColorRgb;
                h = h * 31 + (Alignment == null ? 0 : Alignment.GetHashCode());
                return h;
            }
        }

        // ---- 宽松模式 ----

        [Flags]
        public enum IgnoreDims
        {
            None = 0,
            SizePt = 1,
            Color = 2,
            Bold = 4
        }

        /// <summary>
        /// 同类判定：双锚点分别相等（勾选忽略的维度在两个锚点上都不参与比对）。
        /// 对齐不参与比较（黄金取证：居中主标题与左对齐模块标题需在"忽略字号"下合并为 8 段）；
        /// 斜体仅记录展示不参与比较。
        /// </summary>
        public bool Matches(FormatSignature other, IgnoreDims ignore)
        {
            if (!FontEquals(EastAsiaFont, other.EastAsiaFont)) return false;
            if (!FontEquals(LatinFont, other.LatinFont)) return false;
            if (!FontEquals(LastEastAsiaFont, other.LastEastAsiaFont)) return false;
            if (!FontEquals(LastLatinFont, other.LastLatinFont)) return false;
            if ((ignore & IgnoreDims.SizePt) == 0 &&
                (!SizePtEquals(SizePt, other.SizePt) || !SizePtEquals(LastSizePt, other.LastSizePt))) return false;
            if ((ignore & IgnoreDims.Bold) == 0 &&
                (!BoldEquals(BoldRaw, other.BoldRaw) || !BoldEquals(LastBoldRaw, other.LastBoldRaw))) return false;
            if ((ignore & IgnoreDims.Color) == 0 &&
                (!ColorEquals(ColorRgb, other.ColorRgb) || !ColorEquals(LastColorRgb, other.LastColorRgb))) return false;
            return true;
        }
    }

    /// <summary>签名提取与同类匹配（PRD v2 P0-1）。</summary>
    public static class FormatSignatureReader
    {
        /// <summary>性能剖析累计（控制器单次匹配后读取并清零）。</summary>
        public static long TicksEnumerator, TicksRange, TicksText, TicksProbe, TicksFont, TicksAlign;

        private static long Now() { return System.Diagnostics.Stopwatch.GetTimestamp(); }
        private static long Delta(long start) { return Now() - start; }
        /// <summary>
        /// 提取段落签名。快路径：段落区间整体读（均匀段落 = 首个非空 run 的有效值，典型文档占绝大多数）；
        /// 出现混合哨兵（wdUndefined）时降级为“首个非空字符”逐点探测。空段返回 IsEmpty 签名。
        /// </summary>
        public static FormatSignature Read(Paragraph paragraph)
        {
            if (paragraph == null) return EmptySig();

            Range probe = null;
            try
            {
                long t0 = Now();
                Range whole = paragraph.Range;
                int start = whole.Start;
                TicksRange += Delta(t0);

                long t1 = Now();
                string text = whole.Text ?? "";
                TicksText += Delta(t1);

                int offset = -1;
                int limit = System.Math.Min(text.Length, 24);
                for (int i = 0; i < limit; i++)
                {
                    char ch = text[i];
                    if (ch == '\r' || ch == '\a' || ch == '\v' || ch == ' ' || ch == '\u3000' || ch == '\t') continue;
                    offset = i;
                    break;
                }

                if (offset >= 0)
                {
                    long t2 = Now();
                    probe = whole.Duplicate;
                    probe.Start = start + offset;
                    probe.End = start + offset + 1;
                    TicksProbe += Delta(t2);
                }

                // 装饰性分隔段（如全文 6 个居中 "/"）：可见文本 ≤3 字符且无 CJK、无字母 → 视同空段
                // （黄金取证：正文严格命中恰为 21，分隔段不得计入正文组）
                string visible = text.Trim((char)13, (char)7, (char)11, ' ', (char)0x3000);
                if (visible.Length <= 3)
                {
                    bool hasCjkOrLetter = false;
                    foreach (char ch in visible)
                    {
                        if ((ch >= (char)0x4E00 && ch <= (char)0x9FFF) || (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z'))
                        { hasCjkOrLetter = true; break; }
                    }
                    if (!hasCjkOrLetter) return EmptySig();
                }
            }
            catch { }

            if (probe == null) return EmptySig();

            var sig = new FormatSignature
            {
                SizePt = double.NaN, ColorRgb = -1,
                LastSizePt = double.NaN, LastBoldRaw = WordFormatUtils.UndefinedInt, LastColorRgb = -1,
                Alignment = "?"
            };
            try
            {
                long t3 = Now();
                Font f = probe.Font;
                sig.EastAsiaFont = NormalizeFont(f.NameFarEast);
                sig.LatinFont = NormalizeFont(f.NameAscii);
                sig.SizePt = WordFormatUtils.NormalizeSize(f.Size);
                sig.BoldRaw = f.Bold;
                sig.ItalicRaw = ItalicOfParagraph(probe, f.Italic);
                sig.ColorRgb = WordFormatUtils.WdColorToRgb((int)f.Color);
                ReadLastAnchor(probe, ref sig);   // 尾锚点；单可见字符段回填为首锚点（均匀）
                TicksFont += Delta(t3);
            }
            catch { }

            try
            {
                long t4 = Now();
                // Paragraph.Alignment 经 Paragraph 代理读取实测 ~1.5ms/次；
                // 走探针区间的 ParagraphFormat 同样便宜（与 Font 同路径）
                sig.Alignment = AlignmentKey(Convert.ToInt32(probe.ParagraphFormat.Alignment));
                TicksAlign += Delta(t4);
            }
            catch
            {
                try { sig.Alignment = AlignmentKey(Convert.ToInt32(paragraph.Alignment)); }
                catch { }
            }

            return sig;
        }

        private static FormatSignature EmptySig()
        {
            return new FormatSignature
            {
                SizePt = double.NaN, ColorRgb = -1,
                LastSizePt = double.NaN, LastBoldRaw = WordFormatUtils.UndefinedInt, LastColorRgb = -1,
                Alignment = "?"
            };
        }

        /// <summary>末个可见字符锚点（与首字符不同位置才算；单可见字符段返回 null → 视为均匀段）。</summary>
        private static Range FindLastVisibleCharRange(Range firstCharProbe)
        {
            try
            {
                Range para = firstCharProbe.Paragraphs[1].Range;
                int endPos = para.End - 1;
                int start = para.Start;
                int guard = 0;
                while (endPos > start && guard++ < 6)
                {
                    Range last = para.Document.Range(endPos - 1, endPos);
                    string t = last.Text ?? "";
                    if (t.Length > 0)
                    {
                        char ch = t[0];
                        if (ch != '\r' && ch != '\a' && ch != '\v' && ch != ' ' && ch != '　')
                            return last.Start == firstCharProbe.Start ? null : last;
                    }
                    endPos--;
                }
                return null;
            }
            catch { return null; }
        }

        /// <summary>填充尾锚点；找不到独立的末字符（单可见字符段）时回填首锚点值（均匀语义）。</summary>
        private static void ReadLastAnchor(Range firstCharProbe, ref FormatSignature sig)
        {
            Range lastChar = FindLastVisibleCharRange(firstCharProbe);
            if (lastChar == null)
            {
                sig.LastEastAsiaFont = sig.EastAsiaFont;
                sig.LastLatinFont = sig.LatinFont;
                sig.LastSizePt = sig.SizePt;
                sig.LastBoldRaw = sig.BoldRaw;
                sig.LastColorRgb = sig.ColorRgb;
                return;
            }
            try
            {
                Font lf = lastChar.Font;
                sig.LastEastAsiaFont = NormalizeFont(lf.NameFarEast);
                sig.LastLatinFont = NormalizeFont(lf.NameAscii);
                sig.LastSizePt = WordFormatUtils.NormalizeSize(lf.Size);
                sig.LastBoldRaw = lf.Bold;
                sig.LastColorRgb = WordFormatUtils.WdColorToRgb((int)lf.Color);
            }
            catch
            {
                // 尾锚点读取失败 → 保持未知（该段与谁都不判同类，宁可少命中不可错命中）
            }
        }

        /// <summary>Word 对未知字体名可能返回空串，归一化为 null（与其他未知一起参与全等比较）。</summary>
        private static string NormalizeFont(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return name;
        }

        /// <summary>
        /// 段落级斜体语义：首字符与末个可见字符均为 -1 才算整体斜体段（-1），否则 0。
        /// 混合斜体段（如仅标点斜体）判非斜体——黄金取证：斜体图注与非斜体副标题需可合并为 7 段，
        /// 而整体斜体的正文段不得混入正文组（21/64 计数）。
        /// </summary>
        private static int ItalicOfParagraph(Range firstCharProbe, int firstItalic)
        {
            // 二态（黄金取证）：首字符与末个可见字符均为 -1 才算整体斜体段；其余（含杂值/混合）一律 0。
            // 斜体标点（【 图注）与纯段落（副标题）由此可合并为 7 段；整体斜体差异仍可区分。
            if (firstItalic != -1) return 0;
            try
            {
                Range para = firstCharProbe.Paragraphs[1].Range;
                int endPos = para.End - 1;
                int start = para.Start;
                int guard = 0;
                while (endPos > start && guard++ < 6)
                {
                    Range last = para.Document.Range(endPos - 1, endPos);
                    string t = last.Text ?? "";
                    if (t.Length > 0)
                    {
                        char ch = t[0];
                        if (ch != '\r' && ch != '\a' && ch != '\v' && ch != ' ' && ch != '\u3000')
                        {
                            int it = last.Font.Italic;
                            return it == -1 ? -1 : 0;
                        }
                    }
                    endPos--;
                }
                return 0;
            }
            catch { return 0; }
        }

        /// <summary>ColorRgb 的"混合"标记（观测到段内颜色不一致）。区别于 -1（自动/未知）：
        /// 两个 -1 视为相等（同为自动色），混合标记不与任何值相等。见 FormatSignature.ColorMixed。</summary>
        private static int NormalizeItalic(int v)
        {
            return v == -1 || v == 0 ? v : WordFormatUtils.UndefinedInt;
        }

        private static string AlignmentKey(int wdAlign)
        {
            switch (wdAlign)
            {
                case (int)WdParagraphAlignment.wdAlignParagraphLeft: return "l";
                case (int)WdParagraphAlignment.wdAlignParagraphCenter: return "c";
                case (int)WdParagraphAlignment.wdAlignParagraphRight: return "r";
                case (int)WdParagraphAlignment.wdAlignParagraphJustify:
                case (int)WdParagraphAlignment.wdAlignParagraphDistribute: return "j";
                default: return "?";
            }
        }
    }
}
