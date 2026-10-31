using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formatting;

namespace TypeRight.Signature
{
    /// <summary>
    /// 格式签名（PRD v2 §2）：段落首个非空 run 的有效格式元组
    /// (中文字体, 西文字体, 字号, 加粗, 斜体, 颜色, 对齐)。读取的是 Word 解析后的有效值。
    /// </summary>
    public struct FormatSignature : IEquatable<FormatSignature>
    {
        public string EastAsiaFont;   // null = 未知/混合
        public string LatinFont;
        public double SizePt;         // NaN = 未知/混合
        public int BoldRaw;           // -1 / 0 / 9999999(混合)
        public int ItalicRaw;
        public int ColorRgb;          // -1 = 自动/混合；0xRRGGBB
        public string Alignment;      // "l"/"c"/"r"/"j"/"?"

        /// <summary>ColorRgb 的"混合"标记（段内颜色不一致）。区别于 -1（自动/未知）：
        /// 两个 -1 视为相等（同为自动色），混合标记不与任何值相等。</summary>
        public const int ColorMixed = -2;

        public bool Bold { get { return BoldRaw == -1; } }
        public bool BoldMixed { get { return BoldRaw == WordFormatUtils.UndefinedInt; } }
        public bool Italic { get { return ItalicRaw == -1; } }

        public bool IsEmpty
        {
            get { return EastAsiaFont == null && LatinFont == null && double.IsNaN(SizePt); }
        }

        public override string ToString()
        {
            string colorKey = ColorRgb == ColorMixed ? "混合" : (ColorRgb < 0 ? "自动" : ColorRgb.ToString("X6"));
            return string.Format("{0}/{1} {2}pt B={3} I={4} #{5} {6}",
                EastAsiaFont ?? "?", LatinFont ?? "?",
                double.IsNaN(SizePt) ? "?" : SizePt.ToString("0.#"),
                BoldMixed ? "?" : (Bold ? 1 : 0),
                ItalicRaw,
                colorKey,
                Alignment);
        }

        public bool Equals(FormatSignature other)
        {
            return string.Equals(EastAsiaFont, other.EastAsiaFont, StringComparison.OrdinalIgnoreCase)
                && string.Equals(LatinFont, other.LatinFont, StringComparison.OrdinalIgnoreCase)
                && SizePtEquals(SizePt, other.SizePt)
                && BoldEquals(this, other)
                && ColorEquals(ColorRgb, other.ColorRgb);
        }

        /// <summary>颜色相等：混合标记（ColorMixed）不与任何值相等；两个 -1（自动）相等。</summary>
        private static bool ColorEquals(int a, int b)
        {
            if (a == ColorMixed || b == ColorMixed) return false;
            return a == b;
        }

        // Word 的 Font.Bold/Italic 会返回非标准杂值（实测混入颜色字节，如 1717853）。
        // 归一化：只有 -1 视为真，0 视为假，其余（9999999/杂值）= 未知。
        public bool BoldEffective { get { return BoldRaw == -1; } }
        public bool BoldKnown { get { return BoldRaw == -1 || BoldRaw == 0; } }
        public bool ItalicEffective { get { return ItalicRaw == -1; } }

        // 未知字号不等于未知字号：读取失败的段落曾因 NaN==NaN 判等而互相"同类"（跨判 bug）。
        // 严格方向取舍：宁可少命中，不可错命中。
        private static bool SizePtEquals(double a, double b)
        {
            if (double.IsNaN(a) || double.IsNaN(b)) return false;
            return Math.Abs(a - b) < 0.01;
        }

        // 加粗含未知（杂值）即不同类：杂值曾折为"非加粗"，与真·非加粗跨判
        // （视觉加粗但读到杂值的标题，会与正文判同类）。
        private static bool BoldEquals(FormatSignature a, FormatSignature b)
        {
            if (!a.BoldKnown || !b.BoldKnown) return false;
            return a.BoldRaw == b.BoldRaw;
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
                h = h * 31 + ItalicRaw;
                h = h * 31 + ColorRgb;
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
        /// 宽松比较：勾选忽略的维度不参与比对。
        /// 对齐不参与比较（黄金文件取证：居中主标题与左对齐模块标题需在“忽略字号”下合并为 8 段）；
        /// 斜体/加粗按归一化有效值比较（Word 杂值容忍）。
        /// </summary>
        public bool Matches(FormatSignature other, IgnoreDims ignore)
        {
            if (!string.Equals(EastAsiaFont, other.EastAsiaFont, StringComparison.OrdinalIgnoreCase)) return false;
            if (!string.Equals(LatinFont, other.LatinFont, StringComparison.OrdinalIgnoreCase)) return false;
            if ((ignore & IgnoreDims.SizePt) == 0 && !SizePtEquals(SizePt, other.SizePt)) return false;
            if ((ignore & IgnoreDims.Bold) == 0 && !BoldEquals(this, other)) return false;
            if ((ignore & IgnoreDims.Color) == 0 && !ColorEquals(ColorRgb, other.ColorRgb)) return false;
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

            var sig = new FormatSignature { SizePt = double.NaN, ColorRgb = -1, Alignment = "?" };
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
                // 首尾一致性：行首符号/序号与正文不同格式的"混合段"（如 ★10.5pt + 标题16pt）
                // 曾按首字符整段判类，与均匀段跨判同类 → 首尾不一致的维度置未知（严格模式不判同类）
                // 注意 sig 是 struct，必须 ref 传参（按值改的是副本）
                ApplyUniformityCheck(probe, ref sig);
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
            return new FormatSignature { SizePt = double.NaN, ColorRgb = -1, Alignment = "?" };
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
        private const int ColorMixed = FormatSignature.ColorMixed;

        /// <summary>
        /// 首尾一致性校验：末个可见字符与首字符逐维比对，不一致的维度置"未知/混合"。
        /// 严格模式未知即不同类；宽松模式可勾选忽略对应维度后仍可合并（语义正确）。
        /// 首尾一致但中间不同的情况极罕见（黄金取证未见），按均匀处理。
        /// </summary>
        private static void ApplyUniformityCheck(Range firstCharProbe, ref FormatSignature sig)
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
                        {
                            Font lf = last.Font;
                            if (!double.IsNaN(sig.SizePt))
                            {
                                double lastSize = WordFormatUtils.NormalizeSize(lf.Size);
                                if (!double.IsNaN(lastSize) && Math.Abs(lastSize - sig.SizePt) >= 0.01)
                                    sig.SizePt = double.NaN;                    // 混合字号 → 未知
                            }
                            if (sig.BoldKnown)
                            {
                                int lastBold = lf.Bold;
                                if (lastBold != sig.BoldRaw) sig.BoldRaw = WordFormatUtils.UndefinedInt; // 混合加粗 → 未知
                            }
                            if (sig.ColorRgb >= 0)
                            {
                                int lastColor = WordFormatUtils.WdColorToRgb((int)lf.Color);
                                if (lastColor != sig.ColorRgb) sig.ColorRgb = ColorMixed;   // 混合颜色 → 混合标记
                            }
                            return;
                        }
                    }
                    endPos--;
                }
            }
            catch { }
        }

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
