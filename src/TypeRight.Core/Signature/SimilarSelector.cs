using System;
using System.Collections.Generic;
using System.Diagnostics;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formatting;

namespace TypeRight.Signature
{
    /// <summary>同类匹配结果。</summary>
    public sealed class SimilarMatchResult
    {
        /// <summary>命中段落（文档序）。</summary>
        public List<Paragraph> Paragraphs = new List<Paragraph>();
        public FormatSignature SampleSignature;
        public FormatSignature.IgnoreDims Ignore;
        public TimeSpan Duration;

        public int Count { get { return Paragraphs.Count; } }
    }

    /// <summary>匹配单元（引导句感知）：段内一个连续区间——均匀段=整段；
    /// 混合段拆为"引导句（行首格式异于后续的首个连续区）+正文"两个单元。
    /// ★ 导读框即此类：行首"★ 标题句"加粗蓝 + 后续宋体正文同处一段。</summary>
    public sealed class MatchUnit
    {
        public Paragraph Paragraph;
        /// <summary>单元区间 [Start, End)（文档坐标，不含段落/单元格标记）。</summary>
        public int Start, End;
        /// <summary>是否引导句单元。</summary>
        public bool IsLead;
        public FormatSignature Signature;
    }

    /// <summary>单元匹配结果。</summary>
    public sealed class UnitMatchResult
    {
        public List<MatchUnit> Units = new List<MatchUnit>();
        public FormatSignature SampleSignature;
        public bool SampleIsLead;
        public TimeSpan Duration;
        public int Count { get { return Units.Count; } }
    }

    /// <summary>
    /// 同类段落匹配（PRD v2 P0-1）：全文段落 + 表格单元格内段落，按签名比对。
    /// 分区规则（黄金文件取证）：正文样本只匹配正文段落，表格样本只匹配表格内段落——
    /// 避免正文样本把表格里的同格式装饰段（如 ★ 导读）一并卷入。
    /// 性能目标：1000 段 &lt; 1s。
    /// </summary>
    public static class SimilarSelector
    {
        /// <summary>样本段落驱动：提取签名 + 判定表格分区后匹配。</summary>
        public static SimilarMatchResult Match(Document doc, Paragraph sampleParagraph, FormatSignature.IgnoreDims ignore)
        {
            bool sampleInTable = false;
            try { sampleInTable = IsInTable(doc, sampleParagraph.Range.Start); }
            catch { }
            return Match(doc, FormatSignatureReader.Read(sampleParagraph, MathZoneRanges(doc)), ignore, sampleInTable);
        }

        /// <summary>全文表格区间（[start,end) 列表）。正文/表格分区不再用
        /// Range.Information(wdWithInTable)——该调用在进程内部分状态下会把正文段
        /// 误判为表格段（实测：Ribbon 线程跑过 WPF 渲染后必现，正文段全被跳过 → 0 命中）。
        /// 按段落起点是否落在表格区间内判定，稳定且免去了逐段的 Information 开销。</summary>
        private static List<int[]> TableRanges(Document doc)
        {
            var ranges = new List<int[]>();
            try
            {
                foreach (Table t in WordCompat.TablesOf(doc.Tables))
                {
                    try { ranges.Add(new[] { t.Range.Start, t.Range.End }); } catch { }
                }
            }
            catch { }
            return ranges;
        }

        /// <summary>全文表格区间（供公式转换等模块及验收诊断复用；位置判定替代 wdWithInTable）。</summary>
        public static List<int[]> TableRangesOf(Document doc) { return TableRanges(doc); }

        private static bool IsInTable(List<int[]> tableRanges, int position)
        {
            foreach (int[] r in tableRanges)
                if (position >= r[0] && position < r[1]) return true;
            return false;
        }

        private static bool IsInTable(Document doc, int position)
        {
            return IsInTable(TableRanges(doc), position);
        }

        /// <summary>样本段是否在表格内（调用方保存进会话，匹配时传入 4 参重载）。</summary>
        public static bool IsInTable(Document doc, Paragraph paragraph)
        {
            try { return IsInTable(TableRanges(doc), paragraph.Range.Start); }
            catch { return false; }
        }

        /// <summary>全文公式区（OMML）[start,end) 列表（一次构建，供签名读取器跳过公式区锚点）。
        /// 无公式返回 null（零开销快路径）。</summary>
        public static List<int[]> MathZoneRanges(Document doc)
        {
            try
            {
                int count = doc.OMaths.Count;
                if (count == 0) return null;
                var zones = new List<int[]>(count);
                foreach (var m in WordCompat.OMathsOf(doc.OMaths))
                {
                    try { Range mr = m.Range; zones.Add(new[] { mr.Start, mr.End }); }
                    catch { }
                }
                zones.Sort((a, b) => a[0].CompareTo(b[0]));
                return zones;
            }
            catch { return null; }
        }

        // ───────── 引导句感知单元匹配（方案1） ─────────

        /// <summary>检测段内引导句边界：引导句 = 行首起与首锚格式（字号/加粗/颜色）一致的
        /// 最长前缀，边界 = 首个不一致的可见字符位置（文档坐标）。均匀段/拆不出边界返回 -1。
        /// 字体不参与边界判定（Font.Name* 读取抖动）；公式区字符透明跳过（区内读数为数学字体值）。
        /// 定位用顺序逐字符扫描（上限 600 探针）：二分假设"引导区连续前缀"，
        /// 公式化后的段（zone 跳读 + 段中岛状格式）不满足单调性——二分会越过真实边界
        /// 把正文头几字并进引导单元（引导签名被污染掉队），或整段误判拆不出边界。</summary>
        public static int DetectLeadBoundary(Paragraph p, FormatSignature paraSig, List<int[]> mathZones = null)
        {
            if (paraSig.IsUniform) return -1;
            try
            {
                Range whole = p.Range;
                int paraStart = whole.Start;
                string text = whole.Text ?? "";
                int firstIdx = FormatSignatureReader.FirstVisibleIndex(text, 0, paraStart, mathZones);
                int lastIdx = FormatSignatureReader.LastVisibleIndex(text, paraStart, mathZones);
                if (firstIdx < 0 || lastIdx <= firstIdx) return -1;

                Document doc = whole.Document;
                int pos = firstIdx, probes = 0;
                while (pos < lastIdx && probes++ < 600)
                {
                    // 跳过空白与公式区字符（公式区属正文语流，不当边界判据）
                    int probe = pos, guard = 0;
                    while (probe < lastIdx && guard++ < 16)
                    {
                        if (FormatSignatureReader.InMathZone(mathZones, paraStart + probe))
                        {
                            int nz = FormatSignatureReader.SkipMathForward(mathZones, paraStart + probe) - paraStart;
                            if (nz <= probe) break;
                            probe = nz;
                            continue;
                        }
                        if (!FormatSignatureReader.IsSkippable(text, probe)) break;
                        probe++;
                    }
                    if (probe >= lastIdx) break;
                    if (!CharMatchesLead(doc, paraStart + probe, paraSig))
                        return paraStart + probe;   // 首个非引导格式字符 = 精确边界
                    pos = probe + 1;
                }
                return -1;   // 扫到底皆为引导格式（视作均匀）
            }
            catch { return -1; }
        }

        private static bool CharMatchesLead(Document doc, int pos, FormatSignature lead)
        {
            try
            {
                Font f = doc.Range(pos, pos + 1).Font;
                double size = WordFormatUtils.NormalizeSize(f.Size);
                if (!double.IsNaN(size) && !double.IsNaN(lead.SizePt) && Math.Abs(size - lead.SizePt) >= 0.01) return false;
                int bold = f.Bold == -1 || f.Bold == 0 ? f.Bold : WordFormatUtils.UndefinedInt;
                if (lead.BoldRaw != WordFormatUtils.UndefinedInt && bold != WordFormatUtils.UndefinedInt && bold != lead.BoldRaw) return false;
                int color = WordFormatUtils.WdColorToRgb((int)f.Color);
                if (lead.ColorRgb >= 0 && color >= 0 && color != lead.ColorRgb) return false;
                return true;
            }
            catch { return false; }
        }

        /// <summary>引导句感知全文匹配：sampleOffset = 光标在样本段内的字符偏移，
        /// 决定样本单元是引导句还是正文。均匀段落单元=整段（行为与段落级匹配一致）。
        /// 签名读取全程公式区感知（mathZones 一次构建）：公式化后的"文字+公式"混合段
        /// 锚点只取文字部分，不再被数学字体读数污染（用户实测：公式化后 24 处同类只剩 7 处）。</summary>
        public static UnitMatchResult MatchUnits(Document doc, Paragraph sampleParagraph, int sampleOffset,
            FormatSignature.IgnoreDims ignore)
        {
            var watch = Stopwatch.StartNew();
            var result = new UnitMatchResult();
            if (doc == null || sampleParagraph == null) return result;

            bool sampleInTable = IsInTable(doc, sampleParagraph);
            List<int[]> mathZones = MathZoneRanges(doc);

            // 样本单元与签名
            FormatSignature sampleParaSig = FormatSignatureReader.Read(sampleParagraph, mathZones);
            if (sampleParaSig.IsEmpty) return result;
            int leadBoundary = DetectLeadBoundary(sampleParagraph, sampleParaSig, mathZones);
            int paraStart = sampleParagraph.Range.Start;
            bool sampleIsLead = leadBoundary >= 0 && sampleOffset >= 0 && sampleOffset < leadBoundary - paraStart;
            result.SampleIsLead = sampleIsLead;

            Range sampleUnitRange = UnitRange(doc, sampleParagraph, leadBoundary, sampleIsLead);
            FormatSignature sampleSig = FormatSignatureReader.ReadRange(sampleUnitRange, mathZones);
            if (sampleSig.IsEmpty) return result;
            result.SampleSignature = sampleSig;

            List<int[]> tableRanges = TableRanges(doc);
            foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
            {
                int pStart;
                try { pStart = p.Range.Start; } catch { continue; }
                if (IsInTable(tableRanges, pStart) != sampleInTable) continue;

                FormatSignature sig = FormatSignatureReader.Read(p, mathZones);
                if (sig.IsEmpty) continue;

                if (sig.IsUniform)
                {
                    // 均匀段：单单元=整段；签名即段签名
                    if (sig.Matches(sampleSig, ignore))
                        result.Units.Add(new MatchUnit { Paragraph = p, Start = pStart, End = p.Range.End, IsLead = false, Signature = sig });
                    continue;
                }

                // 混合段：锚点预筛（引导句候选=首锚匹配；正文候选=尾锚匹配），
                // 两头都不匹配的段直接跳过——避免对全文混合段做昂贵的边界二分+双区间签名
                bool leadPossible = sig.FirstAnchorMatches(sampleSig, ignore);
                bool bodyPossible = sig.LastAnchorMatches(sampleSig, ignore);
                if (!leadPossible && !bodyPossible) continue;

                int boundary = DetectLeadBoundary(p, sig, mathZones);
                if (boundary < 0)
                {
                    // 拆不出引导句（如全段皆首锚格式）：按整段处理
                    if (sig.Matches(sampleSig, ignore))
                        result.Units.Add(new MatchUnit { Paragraph = p, Start = pStart, End = p.Range.End, IsLead = false, Signature = sig });
                    continue;
                }
                if (leadPossible)
                {
                    FormatSignature leadSig = FormatSignatureReader.ReadRange(doc.Range(pStart, boundary), mathZones);
                    if (!leadSig.IsEmpty && leadSig.Matches(sampleSig, ignore))
                        result.Units.Add(new MatchUnit { Paragraph = p, Start = pStart, End = boundary, IsLead = true, Signature = leadSig });
                }
                if (bodyPossible)
                {
                    FormatSignature bodySig = FormatSignatureReader.ReadRange(doc.Range(boundary, p.Range.End), mathZones);
                    if (!bodySig.IsEmpty && bodySig.Matches(sampleSig, ignore))
                        result.Units.Add(new MatchUnit { Paragraph = p, Start = boundary, End = p.Range.End, IsLead = false, Signature = bodySig });
                }
            }

            watch.Stop();
            result.Duration = watch.Elapsed;
            return result;
        }

        private static Range UnitRange(Document doc, Paragraph p, int leadBoundary, bool isLead)
        {
            int start = p.Range.Start, end = p.Range.End;
            if (leadBoundary >= 0)
            {
                if (isLead) end = leadBoundary;
                else start = leadBoundary;
            }
            return doc.Range(start, end);
        }

        /// <summary>样本签名 → 全文匹配（分区：sampleInTable 决定命中文本域）。公式区感知。</summary>
        public static SimilarMatchResult Match(Document doc, FormatSignature sample, FormatSignature.IgnoreDims ignore, bool sampleInTable = false)
        {
            var watch = Stopwatch.StartNew();
            var result = new SimilarMatchResult { SampleSignature = sample, Ignore = ignore };
            if (doc == null || sample.IsEmpty) return result;

            List<int[]> tableRanges = TableRanges(doc);
            List<int[]> mathZones = MathZoneRanges(doc);

            Paragraphs paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;
            FormatSignatureReader.TicksRange = FormatSignatureReader.TicksText = FormatSignatureReader.TicksProbe =
                FormatSignatureReader.TicksFont = FormatSignatureReader.TicksAlign = 0;
            long enumStart = Stopwatch.GetTimestamp();
            // 注意：Paragraphs[i] 索引器是 O(n) 重扫（千段文档实测 O(n²)≈8.9s），必须用枚举器顺序遍历
            int index = 0;
            foreach (Paragraph p in paragraphs)
            {
                index++;
                int start = -1;
                try { start = p.Range.Start; } catch { }
                if (start < 0) continue;
                if (IsInTable(tableRanges, start) != sampleInTable) continue;   // 正文/表格分区
                FormatSignature s = FormatSignatureReader.Read(p, mathZones);
                if (s.IsEmpty) continue;          // 空段/无可见文本不参与
                if (s.Matches(sample, ignore)) result.Paragraphs.Add(p);
            }
            count = index;
            FormatSignatureReader.TicksEnumerator = Stopwatch.GetTimestamp() - enumStart
                - FormatSignatureReader.TicksRange - FormatSignatureReader.TicksText
                - FormatSignatureReader.TicksProbe - FormatSignatureReader.TicksFont - FormatSignatureReader.TicksAlign;

            watch.Stop();
            result.Duration = watch.Elapsed;
            return result;
        }

        // ---- 命中呈现：临时高亮（PRD P0-1：不写入正文内容语义，退出即还原）----
        // 只高亮原本“无高亮”的段落，原有高亮的段落保持不动（避免无法精确还原混合高亮）。

        public sealed class HighlightSession
        {
            private readonly Document _doc;
            internal readonly List<int> TouchedStarts = new List<int>();
            internal readonly List<int> TouchedEnds = new List<int>();

            public HighlightSession(Document doc) { _doc = doc; }

            /// <summary>把命中段落标记为黄色（原本无高亮的）。返回本次实际标记的段数。</summary>
            public int Apply(SimilarMatchResult match)
            {
                Clear();
                if (match == null) return 0;
                int applied = 0;
                foreach (Paragraph p in match.Paragraphs)
                {
                    try
                    {
                        Range r = p.Range;
                        WdColorIndex original = r.HighlightColorIndex;
                        if (original == WdColorIndex.wdNoHighlight)
                        {
                            r.HighlightColorIndex = WdColorIndex.wdYellow;
                            TouchedStarts.Add(r.Start);
                            TouchedEnds.Add(r.End);
                            applied++;
                        }
                    }
                    catch { }
                }
                return applied;
            }

            /// <summary>把命中单元区间标记为黄色（引导句感知：只标引导句或只标正文）。
            /// 区间必须排除末尾的段落/单元格标记：跨结束标记设置高亮会被 Word 扩散到整段
            /// （实测：正文单元 [边界, End) 整段变黄，引导区被误染；End-1 后恢复精确）。</summary>
            public int ApplyUnits(UnitMatchResult match)
            {
                Clear();
                if (match == null) return 0;
                int applied = 0;
                foreach (MatchUnit u in match.Units)
                {
                    try
                    {
                        int end = u.End > u.Start + 1 ? u.End - 1 : u.End;   // 排除段落/单元格标记
                        Range r = _doc.Range(u.Start, end);
                        WdColorIndex original = r.HighlightColorIndex;
                        if (original == WdColorIndex.wdNoHighlight)
                        {
                            r.HighlightColorIndex = WdColorIndex.wdYellow;
                            TouchedStarts.Add(r.Start);
                            TouchedEnds.Add(r.End);
                            applied++;
                        }
                    }
                    catch { }
                }
                return applied;
            }

            /// <summary>还原本次会话标记的全部高亮。</summary>
            public void Clear()
            {
                for (int i = 0; i < TouchedStarts.Count; i++)
                {
                    try
                    {
                        Range r = _doc.Range(TouchedStarts[i], TouchedEnds[i]);
                        r.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                    }
                    catch { }
                }
                TouchedStarts.Clear();
                TouchedEnds.Clear();
            }
        }
    }
}
