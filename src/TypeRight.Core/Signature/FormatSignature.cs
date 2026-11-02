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

        /// <summary>中段采样锚点（段落可见区间 25%/50%/75% 三个位置）。
        /// 双锚点对"首尾相同、中间不同"的两个家族无能为力（用户实测命中），
        /// 中段采样补盲区。null（未提供）与 null 相等；null 与非 null 不等（严格）。</summary>
        public struct AnchorExtra
        {
            public string Ea, Latin;
            public double Size;
            public int Bold, Color;
        }

        public AnchorExtra[] ExtraAnchors;

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
            if (IsUniform && MiddleUniform()) return head;
            string tail = IsUniform ? "" : string.Format(" ⇄尾 {0}/{1} {2}pt B={3} #{4}",
                LastEastAsiaFont ?? "?", LastLatinFont ?? "?",
                double.IsNaN(LastSizePt) ? "?" : LastSizePt.ToString("0.#"),
                LastBoldRaw == -1 ? 1 : (LastBoldRaw == 0 ? 0 : -1),
                LastColorRgb < 0 ? "自动" : LastColorRgb.ToString("X6"));
            if (MiddleUniform()) return head + tail;
            var sb = new System.Text.StringBuilder(" ⋯中");
            foreach (FormatSignature.AnchorExtra e in ExtraAnchors)
            {
                sb.Append('[').Append(double.IsNaN(e.Size) ? "?" : e.Size.ToString("0.#"))
                  .Append("pt B").Append(e.Bold == -1 ? 1 : (e.Bold == 0 ? 0 : -1)).Append(']');
            }
            return head + tail + sb;
        }

        /// <summary>中段采样点是否都与首锚一致（公式区污染检测：中段读到公式区内部
        /// 字体值时为 false——正常"文字+公式"段在公式区感知下应为 true）。</summary>
        public bool MiddleUniform()
        {
            if (ExtraAnchors == null) return true;
            foreach (FormatSignature.AnchorExtra e in ExtraAnchors)
            {
                if (!SizePtEquals(e.Size, SizePt) || !BoldEquals(e.Bold, BoldRaw) ||
                    !ColorEquals(e.Color, ColorRgb)) return false;   // 字体不比（读取抖动）
            }
            return true;
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

        /// <summary>首锚部分匹配（引导句候选预筛）：仅比首锚的字体/字号/加粗/颜色。</summary>
        public bool FirstAnchorMatches(FormatSignature other, IgnoreDims ignore)
        {
            if (!FontEquals(EastAsiaFont, other.EastAsiaFont)) return false;
            if (!FontEquals(LatinFont, other.LatinFont)) return false;
            if ((ignore & IgnoreDims.SizePt) == 0 && !SizePtEquals(SizePt, other.SizePt)) return false;
            if ((ignore & IgnoreDims.Bold) == 0 && !BoldEquals(BoldRaw, other.BoldRaw)) return false;
            if ((ignore & IgnoreDims.Color) == 0 && !ColorEquals(ColorRgb, other.ColorRgb)) return false;
            return true;
        }

        /// <summary>尾锚部分匹配（正文候选预筛）：仅比尾锚的字体/字号/加粗/颜色。</summary>
        public bool LastAnchorMatches(FormatSignature other, IgnoreDims ignore)
        {
            if (!FontEquals(LastEastAsiaFont, other.LastEastAsiaFont)) return false;
            if (!FontEquals(LastLatinFont, other.LastLatinFont)) return false;
            if ((ignore & IgnoreDims.SizePt) == 0 && !SizePtEquals(LastSizePt, other.LastSizePt)) return false;
            if ((ignore & IgnoreDims.Bold) == 0 && !BoldEquals(LastBoldRaw, other.LastBoldRaw)) return false;
            if ((ignore & IgnoreDims.Color) == 0 && !ColorEquals(LastColorRgb, other.LastColorRgb)) return false;
            return true;
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
        /// 同类判定：双锚点 + 中段采样锚点分别相等（勾选忽略的维度在所有锚点都不参与比对）。
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

            // 中段采样：首尾相同时，中段不同的两个家族在此分离。
            // 中段不比字体：Font.Name* 对部分区间存在抖动（同字符多次读取返回不同字体名，
            // 实测 MS Mincho ↔ 微软雅黑），采样放大读取量后必然误伤；字号/加粗/颜色足以分离。
            FormatSignature.AnchorExtra[] ea = ExtraAnchors, eb = other.ExtraAnchors;
            if ((ea == null) != (eb == null)) return false;
            if (ea != null)
            {
                if (ea.Length != eb.Length) return false;
                for (int i = 0; i < ea.Length; i++)
                {
                    if ((ignore & IgnoreDims.SizePt) == 0 && !SizePtEquals(ea[i].Size, eb[i].Size)) return false;
                    if ((ignore & IgnoreDims.Bold) == 0 && !BoldEquals(ea[i].Bold, eb[i].Bold)) return false;
                    if ((ignore & IgnoreDims.Color) == 0 && !ColorEquals(ea[i].Color, eb[i].Color)) return false;
                }
            }
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

        /// <summary>任意子区间（引导句/正文单元）的签名：锚点在区间内的首/末可见字符与中段采样。
        /// 与段落级 Read 同一套锚点语义，供"引导句感知"单元匹配使用。
        /// mathZones：文档级公式区 [start,end) 列表——锚点不得落在公式区内
        /// （公式区内部字符的字体读数是数学字体值，会污染签名导致公式化后的段落互相匹配不上）。</summary>
        public static FormatSignature ReadRange(Range range, List<int[]> mathZones = null)
        {
            if (range == null) return EmptySig();
            try
            {
                string text = range.Text ?? "";
                int basePos = range.Start;
                int firstIdx = FirstVisibleIndex(text, 0, basePos, mathZones);
                if (firstIdx < 0) return EmptySig();
                int lastIdx = LastVisibleIndex(text, basePos, mathZones);
                if (lastIdx <= firstIdx) lastIdx = firstIdx;

                var sig = new FormatSignature
                {
                    SizePt = double.NaN, ColorRgb = -1,
                    LastSizePt = double.NaN, LastBoldRaw = WordFormatUtils.UndefinedInt, LastColorRgb = -1,
                    Alignment = "l"
                };
                Document doc = range.Document;

                // 首锚
                FormatSignature.AnchorExtra a = ReadAnchor(doc.Range(basePos + firstIdx, basePos + firstIdx + 1), compareFonts: true);
                sig.EastAsiaFont = a.Ea; sig.LatinFont = a.Latin;
                sig.SizePt = a.Size; sig.BoldRaw = a.Bold; sig.ColorRgb = a.Color;
                sig.ItalicRaw = 0;

                // 尾锚
                if (lastIdx == firstIdx)
                {
                    sig.LastEastAsiaFont = sig.EastAsiaFont; sig.LastLatinFont = sig.LatinFont;
                    sig.LastSizePt = sig.SizePt; sig.LastBoldRaw = sig.BoldRaw; sig.LastColorRgb = sig.ColorRgb;
                }
                else
                {
                    FormatSignature.AnchorExtra la = ReadAnchor(doc.Range(basePos + lastIdx, basePos + lastIdx + 1), compareFonts: true);
                    sig.LastEastAsiaFont = la.Ea; sig.LastLatinFont = la.Latin;
                    sig.LastSizePt = la.Size; sig.LastBoldRaw = la.Bold; sig.LastColorRgb = la.Color;
                }

                // 中段 25/50/75%
                var extras = new FormatSignature.AnchorExtra[3];
                for (int k = 0; k < 3; k++)
                {
                    int idx = firstIdx + (int)Math.Round((lastIdx - firstIdx) * 0.25 * (k + 1));
                    int guard = 0;
                    while (idx < lastIdx && guard++ < 8)
                    {
                        if (InMathZone(mathZones, basePos + idx))
                        {
                            int nz = SkipMathForward(mathZones, basePos + idx) - basePos;
                            if (nz <= idx) break;
                            idx = nz;
                            continue;
                        }
                        if (!IsSkippable(text, idx)) break;
                        idx++;
                    }
                    if (idx <= firstIdx || idx >= lastIdx)
                        extras[k] = FirstAnchorOf(sig);
                    else
                        extras[k] = ReadAnchor(doc.Range(basePos + idx, basePos + idx + 1), compareFonts: false);
                }
                sig.ExtraAnchors = extras;
                return sig;
            }
            catch { return EmptySig(); }
        }

        // ---- 公式区感知（mathZones 为文档级 OMML 区间 [start,end)，排序后传入；null=全文无公式）----

        /// <summary>文档位置是否在公式区内。</summary>
        internal static bool InMathZone(List<int[]> zones, int pos)
        {
            if (zones == null) return false;
            for (int i = 0; i < zones.Count; i++)
            {
                int[] z = zones[i];
                if (pos >= z[0] && pos < z[1]) return true;
                if (z[0] > pos) break;   // 已排序，后面更远
            }
            return false;
        }

        /// <summary>把位置推进到公式区之后（连续多区跳完为止）。</summary>
        internal static int SkipMathForward(List<int[]> zones, int pos)
        {
            if (zones == null) return pos;
            bool moved = true;
            int guard = 0;
            while (moved && guard++ < 16)
            {
                moved = false;
                for (int i = 0; i < zones.Count; i++)
                {
                    int[] z = zones[i];
                    if (pos >= z[0] && pos < z[1]) { pos = z[1]; moved = true; break; }
                    if (z[0] > pos) break;
                }
            }
            return pos;
        }

        internal static bool IsSkippable(string text, int idx)
        {
            if (idx < 0 || idx >= text.Length) return true;
            char ch = text[idx];
            return ch == '\r' || ch == '\a' || ch == '\v' || ch == ' ' || ch == '　' || ch == '\t';
        }

        internal static int FirstVisibleIndex(string text, int from)
        {
            for (int i = from; i < text.Length; i++)
                if (!IsSkippable(text, i)) return i;
            return -1;
        }

        internal static int LastVisibleIndex(string text)
        {
            for (int i = text.Length - 1; i >= 0; i--)
                if (!IsSkippable(text, i)) return i;
            return -1;
        }

        /// <summary>公式区感知版：跳过空白与公式区字符，返回首个可充当锚点的字符下标。</summary>
        internal static int FirstVisibleIndex(string text, int from, int basePos, List<int[]> zones)
        {
            for (int i = from; i < text.Length; i++)
            {
                if (IsSkippable(text, i)) continue;
                if (InMathZone(zones, basePos + i)) continue;
                return i;
            }
            return -1;
        }

        /// <summary>公式区感知版：返回末个锚点字符下标。</summary>
        internal static int LastVisibleIndex(string text, int basePos, List<int[]> zones)
        {
            for (int i = text.Length - 1; i >= 0; i--)
            {
                if (IsSkippable(text, i)) continue;
                if (InMathZone(zones, basePos + i)) continue;
                return i;
            }
            return -1;
        }

        /// <summary>
        /// 提取段落签名。快路径：段落区间整体读（均匀段落 = 首个非空 run 的有效值，典型文档占绝大多数）；
        /// 出现混合哨兵（wdUndefined）时降级为“首个非空字符”逐点探测。空段返回 IsEmpty 签名。
        /// </summary>
        public static FormatSignature Read(Paragraph paragraph, List<int[]> mathZones = null)
        {
            if (paragraph == null) return EmptySig();

            Range probe = null;
            int start = 0;
            string text = "";
            int offset = -1;
            try
            {
                long t0 = Now();
                Range whole = paragraph.Range;
                start = whole.Start;
                TicksRange += Delta(t0);

                long t1 = Now();
                text = whole.Text ?? "";
                TicksText += Delta(t1);

                offset = -1;
                // \u6709\u516c\u5f0f\u533a\u65f6\u626b\u5168\u6587\uff08\u516c\u5f0f\u53ef\u80fd\u5360\u6bb5\u9996\u5f88\u957f\uff0c24 \u5b57\u7b26\u7a97\u53e3\u4f1a\u6f0f\u6389\u540e\u9762\u7684\u6b63\u6587\u951a\u70b9\uff09
                int limit = mathZones != null ? text.Length : System.Math.Min(text.Length, 24);
                for (int i = 0; i < limit; i++)
                {
                    char ch = text[i];
                    if (ch == '\r' || ch == '\a' || ch == '\v' || ch == ' ' || ch == '\u3000' || ch == '\t') continue;
                    if (InMathZone(mathZones, start + i)) continue;   // \u516c\u5f0f\u533a\u5b57\u7b26\u4e0d\u5145\u5f53\u951a\u70b9
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
                sig.EastAsiaFont = NormalizeFont(f.NameFarEast);   // 首锚经 Duplicate 路径读取，实测稳定，单读
                sig.LatinFont = NormalizeFont(f.NameAscii);
                sig.SizePt = ReadSizeStable(f);
                sig.BoldRaw = ReadBoldStable(f);
                sig.ItalicRaw = ItalicOfParagraph(probe, f.Italic, mathZones);
                sig.ColorRgb = WordFormatUtils.WdColorToRgb((int)f.Color);
                Range lastChar = FindLastVisibleCharRange(probe, mathZones);
                ReadLastAnchor(lastChar, ref sig);        // 尾锚点；单可见字符段回填为首锚点（均匀）
                ReadExtraAnchors(probe, lastChar, start, text, offset, ref sig, mathZones);   // 中段 25/50/75% 采样
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

        /// <summary>末个可见字符锚点（与首字符不同位置才算；单可见字符段返回 null → 视为均匀段）。
        /// 公式区感知：公式区字符不当尾锚，位置落在区内时跳到区首之前继续找。</summary>
        private static Range FindLastVisibleCharRange(Range firstCharProbe, List<int[]> mathZones)
        {
            try
            {
                Range para = firstCharProbe.Paragraphs[1].Range;
                int endPos = para.End - 1;
                int start = para.Start;
                int guard = 0;
                while (endPos > start && guard++ < 64)
                {
                    if (InMathZone(mathZones, endPos - 1))
                    {
                        // 跳过整个公式区：退到区首（区首位置上的字符是区内第一个字符）
                        for (int i = 0; i < (mathZones == null ? 0 : mathZones.Count); i++)
                        {
                            int[] z = mathZones[i];
                            if (endPos - 1 >= z[0] && endPos - 1 < z[1]) { endPos = z[0]; break; }
                            if (z[0] > endPos - 1) break;
                        }
                        continue;
                    }
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
        private static void ReadLastAnchor(Range lastChar, ref FormatSignature sig)
        {
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
                FormatSignature.AnchorExtra a = ReadAnchor(lastChar, compareFonts: true);
                sig.LastEastAsiaFont = a.Ea;
                sig.LastLatinFont = a.Latin;
                sig.LastSizePt = a.Size;
                sig.LastBoldRaw = a.Bold;
                sig.LastColorRgb = a.Color;
            }
            catch
            {
                // 尾锚点读取失败 → 保持未知（该段与谁都不判同类，宁可少命中不可错命中）
            }
        }

        private static FormatSignature.AnchorExtra ReadAnchor(Range oneChar, bool compareFonts)
        {
            Font lf = oneChar.Font;
            return new FormatSignature.AnchorExtra
            {
                Ea = compareFonts ? ReadFontStable(() => lf.NameFarEast) : NormalizeFont(lf.NameFarEast),
                Latin = compareFonts ? ReadFontStable(() => lf.NameAscii) : NormalizeFont(lf.NameAscii),
                Size = ReadSizeStable(lf),
                Bold = ReadBoldStable(lf),
                Color = WordFormatUtils.WdColorToRgb((int)lf.Color)
            };
        }

        /// <summary>字体名读取抖动（同字符多次读取返回不同名，多见于绝对区间探测的尾锚）：
        /// 读两次一致即取；不一致读第三次取多数。仅用于参与字体比较的尾锚。</summary>
        private static string ReadFontStable(Func<string> read)
        {
            string a = read();
            string b = read();
            if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return NormalizeFont(a);
            string c = read();
            if (string.Equals(a, c, StringComparison.OrdinalIgnoreCase)) return NormalizeFont(a);
            if (string.Equals(b, c, StringComparison.OrdinalIgnoreCase)) return NormalizeFont(b);
            return NormalizeFont(b);   // 三次各异：无法判定，取第二次
        }

        /// <summary>Font.Bold 存在间歇性垃圾值（实测混入颜色字节，同一字符多次读取结果不同）。
        /// 重读最多 3 次取有效值（-1/0）；仍垃圾则保留（严格语义：未知不与任何值相等）。</summary>
        private static int ReadBoldStable(Font f)
        {
            int v = f.Bold;
            if (v == -1 || v == 0) return v;
            for (int i = 0; i < 3; i++)
            {
                v = f.Bold;
                if (v == -1 || v == 0) return v;
            }
            return v;
        }

        /// <summary>字号同理重读（wdUndefined 间歇返回）。</summary>
        private static double ReadSizeStable(Font f)
        {
            double v = WordFormatUtils.NormalizeSize(f.Size);
            if (!double.IsNaN(v)) return v;
            for (int i = 0; i < 2; i++)
            {
                v = WordFormatUtils.NormalizeSize(f.Size);
                if (!double.IsNaN(v)) return v;
            }
            return v;
        }

        /// <summary>中段采样锚点：可见区间 25%/50%/75% 三个位置各读一锚。
        /// 双锚点对"首尾相同、中间不同"的两个家族无区分力（用户实测命中），中段采样补盲。
        /// 单可见字符段/极短段：采样点全部回退为首锚点（均匀语义）。</summary>
        private static void ReadExtraAnchors(Range firstCharProbe, Range lastChar, int paraStart,
            string text, int firstIdx, ref FormatSignature sig, List<int[]> mathZones)
        {
            var extras = new FormatSignature.AnchorExtra[3];
            try
            {
                int lastIdx = firstIdx;   // 无独立尾字符 → 全部回退首锚
                if (lastChar != null) lastIdx = lastChar.Start - paraStart;
                for (int k = 0; k < 3; k++)
                {
                    double frac = 0.25 * (k + 1);
                    int idx = firstIdx + (int)Math.Round((lastIdx - firstIdx) * frac);
                    // 落在空白/公式区字符上则向后找锚点字符（不越过尾字符）
                    int guard = 0;
                    while (idx < lastIdx && guard++ < 8)
                    {
                        if (InMathZone(mathZones, paraStart + idx))
                        {
                            int nz = SkipMathForward(mathZones, paraStart + idx) - paraStart;
                            if (nz <= idx) break;
                            idx = nz;
                            continue;
                        }
                        char ch = idx >= 0 && idx < text.Length ? text[idx] : '\0';
                        if (ch != '\r' && ch != '\a' && ch != '\v' && ch != ' ' && ch != '　' && ch != '\t' && ch != '\0') break;
                        idx++;
                    }
                    if (idx <= firstIdx || idx >= lastIdx || idx >= text.Length)
                    {
                        extras[k] = FirstAnchorOf(sig);
                        continue;
                    }
                    Range r = firstCharProbe.Document.Range(paraStart + idx, paraStart + idx + 1);
                    extras[k] = ReadAnchor(r, compareFonts: false);
                }
            }
            catch
            {
                // 读取失败 → 回退首锚（保持段落可比性，中段差异漏检概率极小）
                for (int k = 0; k < 3; k++) extras[k] = FirstAnchorOf(sig);
            }
            sig.ExtraAnchors = extras;
        }

        private static FormatSignature.AnchorExtra FirstAnchorOf(FormatSignature sig)
        {
            return new FormatSignature.AnchorExtra
            {
                Ea = sig.EastAsiaFont, Latin = sig.LatinFont,
                Size = sig.SizePt, Bold = sig.BoldRaw, Color = sig.ColorRgb
            };
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
        private static int ItalicOfParagraph(Range firstCharProbe, int firstItalic, List<int[]> mathZones)
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
                while (endPos > start && guard++ < 64)
                {
                    if (InMathZone(mathZones, endPos - 1))
                    {
                        for (int i = 0; i < (mathZones == null ? 0 : mathZones.Count); i++)
                        {
                            int[] z = mathZones[i];
                            if (endPos - 1 >= z[0] && endPos - 1 < z[1]) { endPos = z[0]; break; }
                            if (z[0] > endPos - 1) break;
                        }
                        continue;
                    }
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
