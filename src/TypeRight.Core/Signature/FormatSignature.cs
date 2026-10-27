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

        public bool Bold { get { return BoldRaw == -1; } }
        public bool BoldMixed { get { return BoldRaw == WordFormatUtils.UndefinedInt; } }
        public bool Italic { get { return ItalicRaw == -1; } }

        public bool IsEmpty
        {
            get { return EastAsiaFont == null && LatinFont == null && double.IsNaN(SizePt); }
        }

        public override string ToString()
        {
            return string.Format("{0}/{1} {2}pt B={3} I={4} #{5:X6} {6}",
                EastAsiaFont ?? "?", LatinFont ?? "?",
                double.IsNaN(SizePt) ? "?" : SizePt.ToString("0.#"),
                BoldMixed ? "?" : (Bold ? 1 : 0),
                ColorRgb < 0 ? -1 : ColorRgb,
                ColorRgb < 0 ? 0 : ColorRgb,
                Alignment);
        }

        public bool Equals(FormatSignature other)
        {
            return string.Equals(EastAsiaFont, other.EastAsiaFont, StringComparison.OrdinalIgnoreCase)
                && string.Equals(LatinFont, other.LatinFont, StringComparison.OrdinalIgnoreCase)
                && SizePtEquals(SizePt, other.SizePt)
                && BoldRaw == other.BoldRaw
                && ItalicRaw == other.ItalicRaw
                && ColorRgb == other.ColorRgb
                && string.Equals(Alignment, other.Alignment, StringComparison.Ordinal);
        }

        private static bool SizePtEquals(double a, double b)
        {
            if (double.IsNaN(a) && double.IsNaN(b)) return true;
            if (double.IsNaN(a) || double.IsNaN(b)) return false;
            return Math.Abs(a - b) < 0.01;
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

        /// <summary>宽松比较：勾选忽略的维度不参与比对（PRD §4 P0-1）。</summary>
        public bool Matches(FormatSignature other, IgnoreDims ignore)
        {
            if (!string.Equals(EastAsiaFont, other.EastAsiaFont, StringComparison.OrdinalIgnoreCase)) return false;
            if (!string.Equals(LatinFont, other.LatinFont, StringComparison.OrdinalIgnoreCase)) return false;
            if (!string.Equals(Alignment, other.Alignment, StringComparison.Ordinal)) return false;
            if (ItalicRaw != other.ItalicRaw) return false;
            if ((ignore & IgnoreDims.SizePt) == 0 && !SizePtEquals(SizePt, other.SizePt)) return false;
            if ((ignore & IgnoreDims.Bold) == 0 && BoldRaw != other.BoldRaw) return false;
            if ((ignore & IgnoreDims.Color) == 0 && ColorRgb != other.ColorRgb) return false;
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
                sig.ItalicRaw = NormalizeItalic(f.Italic);
                sig.ColorRgb = WordFormatUtils.WdColorToRgb((int)f.Color);
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
