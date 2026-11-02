using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Xsl;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formulas;
using TypeRight.Scanning;

namespace TypeRight.Formatting
{
    /// <summary>公式转换统计（PRD v2 §6.4：已转换 N 处，失败 M 处 + 失败定位）。</summary>
    public sealed class FormulaStats
    {
        public int Converted;
        public int Failed;
        /// <summary>整程耗时（毫秒）——性能回归观测用。</summary>
        public long ElapsedMs;
        /// <summary>分段耗时分解（片段扫描/预备/草稿/复制/表格rPr）。</summary>
        public string TimingText;
        public readonly List<string> FailedLocations = new List<string>();

        public string SummaryText()
        {
            return "已转换 " + Converted + " 处，失败 " + Failed + " 处" + (Failed > 0 ? "（失败片段已原样保留）" : "")
                + (ElapsedMs > 0 ? "，" + (ElapsedMs / 1000.0).ToString("0.0") + "s" : "");
        }
    }

    /// <summary>
    /// 纯文本公式 → OMath（PRD v2 §6，规格不变）：切分 → UnicodeMath 线性串 → OMaths.Add → BuildUp。
    /// 失败片段原样保留并计入统计，绝不留半转换残骸。
    /// </summary>
    public sealed class FormulaConverter
    {
        private readonly Document _doc;
        private readonly double? _mathSizePt;
        private readonly string _mml2OmmlPath;

        /// <summary>禁用二遍清扫（诊断用）。</summary>
        public static bool DisableSweep;

        /// <summary>回退记录（验收调试用）。</summary>
        public static readonly System.Collections.Generic.List<string> FallbackLog = new System.Collections.Generic.List<string>();

        private static void RecordFallback(string reason)
        {
            if (FallbackLog.Count < 60) FallbackLog.Add(reason);
        }

        /// <summary>结构构建顺序（验收调试矩阵）：0=Rad→Frac→BuildUp（可保存/脚本缺）；
        /// 1=BuildUp→Rad→Frac（结构全/保存崩）；2=Rad→BuildUp→Frac。</summary>
        public static int StructureOrderMode;

        public FormulaConverter(Document doc, double? mathSizePt = null, string mml2OmmlXslPath = null)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            _doc = doc;
            _mathSizePt = mathSizePt;
            _mml2OmmlPath = mml2OmmlXslPath;
        }

        /// <summary>定位 Office 自带 MML2OMML.XSL（兜底管线用，PRD §6.2.1）。</summary>
        public static string ResolveMml2OmmlPath(Application app)
        {
            var candidates = new List<string>();
            try
            {
                string officeDir = app.Path;
                if (!string.IsNullOrEmpty(officeDir)) candidates.Add(Path.Combine(officeDir, "MML2OMML.XSL"));
            }
            catch { }
            string progFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string progFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            foreach (string root in new[] { progFiles, progFilesX86 })
            {
                if (string.IsNullOrEmpty(root)) continue;
                candidates.Add(Path.Combine(root, @"Microsoft Office\root\Office16\MML2OMML.XSL"));
                candidates.Add(Path.Combine(root, @"Microsoft Office\Office16\MML2OMML.XSL"));
            }
            foreach (string c in candidates)
                if (File.Exists(c)) return c;
            return null;
        }

        /// <summary>转换整个文档（正文段落 + 表格单元格；已是公式对象的区间自动跳过，保证幂等）。
        /// 性能/观感：全程复用一个隐藏草稿文档注入公式（旧行为每个公式 Documents.Open 一次临时文档，
        /// 254 处公式 = 254 次窗口开合：顶部频闪 + ~150ms/个的开销）；
        /// gap 内多片段走批量注入（一个 docx 一次 InsertFile 承载全部，降序逐个复制），
        /// 不再"转一个重扫一遍"（降序转换不影响更早片段的位置）。</summary>
        public FormulaStats ConvertDocument()
        {
            var stats = new FormulaStats();
            var watch = System.Diagnostics.Stopwatch.StartNew();
            ResetTimers();

            // 表格区间（位置判定，替代逐段 wdWithInTable——该调用进程内偶发假值且逐段开销大）
            List<int[]> tableRanges = TypeRight.Signature.SimilarSelector.TableRangesOf(_doc);
            bool hasMml = !string.IsNullOrEmpty(_mml2OmmlPath);

            // 隐藏草稿文档（Visible=false，不生成窗口）：批量注入的公式载体
            Document scratch = null;
            try
            {
                try
                {
                    scratch = _doc.Application.Documents.Add(
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, false);
                }
                catch { scratch = null; }

                foreach (Paragraph p in _doc.Paragraphs)
                {
                    bool convertedHere = false;
                    try
                    {
                        // 单遍降序：每 gap 的片段集合一次算齐，从后往前转换——
                        // 每次插入只影响其后位置（已处理完），更早片段位置稳定，无需"转一个重扫一遍"。
                        // （旧队列式循环：254 处公式 ≈ 100+ 次全段重扫，是"一个个慢慢来"观感的一半来源）
                        foreach (Range gap in TextGapsExcludingMath(p))
                        {
                                long tg = Ts();
                                string gapText;
                                int gapStart;
                                try { gapText = gap.Text ?? ""; gapStart = gap.Start; }
                                catch { continue; }
                                string trimmed = gapText.TrimEnd('\r', '\a', '\v');
                                if (trimmed.Trim().Length == 0) continue;

                                var segs = MathSegmenter.FindSegments(trimmed);
                                MsSeg += Td(tg);
                                if (segs.Count == 0) continue;

                                var done = new bool[segs.Count];

                                // ── 批量注入：gap 内 ≥2 个可准备片段 → 一个 docx 一次装草稿 ──
                                if (hasMml && scratch != null && segs.Count >= 2)
                                {
                                    long tp = Ts();
                                    var mmls = new List<string>();
                                    var order = new List<int>();
                                    for (int si = 0; si < segs.Count; si++)
                                    {
                                        string mml = PrepareMml(segs[si].Text);
                                        if (mml != null) { mmls.Add(mml); order.Add(si); }
                                    }
                                    MsPrepare += Td(tp);   // 只计 MML 预备；注入由 MathmlInjector 自记（草稿/复制）
                                    if (mmls.Count >= 2)
                                    {
                                        var targets = new Range[mmls.Count];
                                        for (int j = 0; j < mmls.Count; j++)
                                        {
                                            var seg = segs[order[j]];
                                            targets[j] = _doc.Range(gapStart + seg.Start, gapStart + seg.Start + seg.Length);
                                        }
                                        bool[] ok = MathmlInjector.TryInjectBatch(_doc, targets, mmls.ToArray(), _mml2OmmlPath, scratch);
                                        for (int j = 0; j < ok.Length; j++)
                                        {
                                            if (!ok[j]) continue;
                                            done[order[j]] = true;
                                            stats.Converted++;
                                            convertedHere = true;
                                        }
                                    }
                                }

                                // ── 降序单发：批量未覆盖/失败/无批量的片段 ──
                                for (int si = segs.Count - 1; si >= 0; si--)
                                {
                                    if (done[si]) continue;
                                    var seg = segs[si];
                                    try
                                    {
                                        long tp = Ts();
                                        long scratchBefore = MathmlInjector.MsScratch, copyBefore = MathmlInjector.MsCopy;
                                        Range rng = _doc.Range(gapStart + seg.Start, gapStart + seg.Start + seg.Length);
                                        bool okSingle = TryConvertRange(rng, seg.Text, false, seg.Start, seg.Length, trimmed.Length, scratch);
                                        MsPrepare += Td(tp) - (MathmlInjector.MsScratch - scratchBefore) - (MathmlInjector.MsCopy - copyBefore);
                                        if (okSingle)
                                        {
                                            stats.Converted++;
                                            convertedHere = true;
                                        }
                                        else
                                        {
                                            stats.Failed++;
                                            if (stats.FailedLocations.Count < 50)
                                                stats.FailedLocations.Add("¶" + ParagraphIndex(p) + "  " + (seg.Text.Length > 30 ? seg.Text.Substring(0, 30) + "…" : seg.Text));
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    catch { }
                    // TC-20（P0）：公式构建会重写单元格 run 的 rPr——只对"表格内且确有公式"的段
                    // 强制显式回写（旧行为对全部表格段都做，纯表格段是无效开销）
                    try
                    {
                        if (convertedHere || HasMathZone(p))
                        {
                            int ps = p.Range.Start;
                            if (InRanges(tableRanges, ps))
                            {
                                long tf = Ts();
                                ForceExplicitRunProps(p);
                                MsForce += Td(tf);
                            }
                        }
                    }
                    catch { }
                }
            }
            finally
            {
                if (scratch != null) { try { scratch.Close(false); } catch { } }
            }
            watch.Stop();
            stats.ElapsedMs = watch.ElapsedMilliseconds;
            stats.TimingText = TimingText(watch.ElapsedMilliseconds);
            return stats;
        }

        // ---- 分段计时（性能归因）----
        public static long MsSeg, MsPrepare, MsForce;
        private static long Ts() { return System.Diagnostics.Stopwatch.GetTimestamp(); }
        private static long Td(long start) { return (System.Diagnostics.Stopwatch.GetTimestamp() - start) * 1000 / System.Diagnostics.Stopwatch.Frequency; }

        private static void ResetTimers()
        {
            MsSeg = MsPrepare = MsForce = 0;
            MathmlInjector.MsScratch = MathmlInjector.MsCopy = 0;
        }

        private static string TimingText(long totalMs)
        {
            long scratch = MathmlInjector.MsScratch, copy = MathmlInjector.MsCopy;
            long other = totalMs - MsSeg - MsPrepare - MsForce - scratch - copy;
            return string.Format("耗时分解: 片段扫描{0}ms 预备(MML/XSL){1}ms 草稿装载{2}ms 公式复制{3}ms 表格rPr{4}ms 其他{5}ms / 总{6}ms",
                MsSeg, MsPrepare, scratch, copy, MsForce, other, totalMs);
        }

        private static bool HasMathZone(Paragraph p)
        {
            try { return p.Range.OMaths.Count > 0; } catch { return false; }
        }

        private static bool InRanges(List<int[]> ranges, int pos)
        {
            if (ranges == null) return false;
            foreach (int[] r in ranges)
                if (pos >= r[0] && pos < r[1]) return true;
            return false;
        }


        /// <summary>失败片段之后：把该 gap 内更早的片段按降序继续转换（失败片段保留）。</summary>
        private void ConvertEarlierSegments(Range gap, List<MathSegmenter.Segment> segs, MathSegmenter.Segment failed,
            Paragraph p, FormulaStats stats)
        {
            for (int i = segs.Count - 2; i >= 0; i--)
            {
                var seg = segs[i];
                if (seg.Start + seg.Length <= failed.Start)
                {
                    try
                    {
                        Range rng = gap.Duplicate;
                        rng.Start = gap.Start + seg.Start;
                        rng.End = rng.Start + seg.Length;
                        if (TryConvertRange(rng, seg.Text, false, seg.Start, seg.Length, gap.Text.Length))
                            stats.Converted++;
                    }
                    catch { }
                }
            }
        }

        /// <summary>段内数学区之外的纯文本子区间（含可见文本）。</summary>
        private static List<Range> TextGapsExcludingMath(Paragraph p)
        {
            var result = new List<Range>();
            Range whole = p.Range;
            int mathCount;
            try { mathCount = whole.OMaths.Count; } catch { mathCount = 0; }
            if (mathCount == 0)
            {
                result.Add(whole);
                return result;
            }
            var bounds = new List<KeyValuePair<int, int>>();
            for (int i = 1; i <= mathCount; i++)
            {
                try
                {
                    Range mr = whole.OMaths[i].Range;
                    bounds.Add(new KeyValuePair<int, int>(mr.Start, mr.End));
                }
                catch { }
            }
            bounds.Sort((a, b) => a.Key.CompareTo(b.Key));
            int start = whole.Start;
            foreach (var b in bounds)
            {
                if (b.Key > start)
                {
                    Range gap = whole.Duplicate;
                    gap.Start = start;
                    gap.End = b.Key;
                    if (HasVisible(gap)) result.Add(gap);
                }
                if (b.Value > start) start = b.Value;
            }
            if (whole.End > start)
            {
                Range tail = whole.Duplicate;
                tail.Start = start;
                tail.End = whole.End;
                if (HasVisible(tail)) result.Add(tail);
            }
            return result;
        }

        /// <summary>转换单个段落内的公式片段。wholeLine=true 时转换整段（独立公式行场景）。
        /// 降序单遍：片段集合从原始文本一次计算，从后往前处理——每次插入只影响其后位置。
        /// （升序重扫方案已废弃：公式区线性文本与底层位置存在错位，公式区后的真实片段会被误杀。）</summary>
        public void ConvertParagraph(Paragraph paragraph, FormulaStats stats, bool wholeLine = false, bool secondPass = false)
        {
            if (paragraph == null || stats == null) return;

            Range baseRange = paragraph.Range;
            string liveText;
            try { liveText = baseRange.Text ?? ""; }
            catch { return; }

            try { if (baseRange.Fields.Count > 0) return; } catch { }

            string trimmed = liveText.TrimEnd('\r', '\a', '\v');
            if (trimmed.Trim().Length == 0) return;

            bool inTable = false;
            try { inTable = Convert.ToBoolean(baseRange.get_Information(WdInformation.wdWithInTable)); }
            catch { }

            List<MathSegmenter.Segment> segments = FindParagraphSegments(liveText, trimmed, wholeLine, secondPass);
            if (segments.Count == 0) return;

            for (int i = segments.Count - 1; i >= 0; i--)
            {
                var seg = segments[i];
                Range rng;
                try
                {
                    rng = baseRange.Duplicate;
                    rng.Start = baseRange.Start + seg.Start;
                    rng.End = rng.Start + seg.Length;
                }
                catch { continue; }

                try { if (rng.OMaths.Count > 0) continue; } catch { }

                // 二遍交集守卫：区间不得与段内任何真实公式区相交。
                // （公式区线性文本经字符串偏移映射会漂移到区外，OMaths 检查失效导致重复造区——
                //   以 COM 级真实区间位置做交集判定，公式区伪片段全被拒绝）
                if (secondPass)
                {
                    bool intersects = false;
                    try
                    {
                        int zoneCount = baseRange.OMaths.Count;
                        for (int z = 1; z <= zoneCount; z++)
                        {
                            Range zr = baseRange.OMaths[z].Range;
                            if (rng.Start < zr.End && zr.Start < rng.End) { intersects = true; break; }
                        }
                    }
                    catch { }
                    if (intersects) continue;
                }

                if (TryConvertRange(rng, seg.Text, secondPass, seg.Start, seg.Length, trimmed.Length))
                    stats.Converted++;
                else
                {
                    stats.Failed++;
                    if (stats.FailedLocations.Count < 50)
                        stats.FailedLocations.Add("¶" + ParagraphIndex(paragraph) + "  " + (seg.Text.Length > 30 ? seg.Text.Substring(0, 30) + "…" : seg.Text));
                }
            }

            // TC-20（P0）：公式构建重写单元格 run 的 rPr，对表格段非公式 run 强制显式回写
            if (inTable) ForceExplicitRunProps(paragraph);
        }

        private static List<MathSegmenter.Segment> FindParagraphSegments(string liveText, string trimmed, bool wholeLine, bool secondPass)
        {
            if (wholeLine)
            {
                int lead = liveText.Length - liveText.TrimStart(' ').Length;
                if (MathSegmenter.IsFormulaCandidate(trimmed.Trim()))
                    return new List<MathSegmenter.Segment>
                    {
                        new MathSegmenter.Segment { Start = lead, Length = trimmed.Trim().Length, Text = trimmed.Trim() }
                    };
            }
            var segments = MathSegmenter.FindSegments(trimmed);
            if (!secondPass) return segments;

            // 二遍（残留清理）：只处理仍含 Unicode 上下标/根号字符的片段——
            // 这些只可能来自未转换的纯文本（已注入公式区的线性文本不含这些字符，映射前已被替换）
            var residueChars = MathSegmenter.SuperscriptChars + MathSegmenter.SubscriptChars + "√";
            var result = new List<MathSegmenter.Segment>();
            foreach (var seg in segments)
            {
                foreach (char ch in seg.Text)
                {
                    if (residueChars.IndexOf(ch) >= 0)
                    {
                        result.Add(seg);
                        break;
                    }
                }
            }
            return result;
        }

        /// <summary>二遍清扫：区间指针漂移（粘贴/撤销副作用）可能残留未转的纯文本片段，补转。</summary>

        /// <summary>删除非表格内、无公式、无图片的空段落（InsertFile 拆段合并的漏网空段）。</summary>
        private void RemoveEmptySideEffectParagraphs()
        {
            for (int pass = 0; pass < 3; pass++)
            {
                bool removed = false;
                var all = new List<Paragraph>();
                foreach (Paragraph p in _doc.Paragraphs) all.Add(p);
                foreach (Paragraph p in all)
                {
                    try
                    {
                        string t = p.Range.Text ?? "";
                        if (t.Trim('\r', '\a', '\v', ' ', '\u3000').Length > 0) continue;
                        try { if (p.Range.OMaths.Count > 0) continue; } catch { }
                        try { if (p.Range.InlineShapes.Count > 0) continue; } catch { }
                        bool inTable;
                        try { inTable = Convert.ToBoolean(p.Range.get_Information(WdInformation.wdWithInTable)); }
                        catch { inTable = true; }
                        if (inTable) continue;
                        p.Range.Delete();
                        removed = true;
                        break;   // 集合已变，重新收集
                    }
                    catch { }
                }
                if (!removed) break;
            }
        }

        public void SweepLeftovers(FormulaStats stats)
        {
            var all = new List<Paragraph>();
            foreach (Paragraph p in _doc.Paragraphs) all.Add(p);
            foreach (Paragraph p in all)
            {
                ConvertParagraph(p, stats, secondPass: true);
            }
        }

        /// <summary>对段落内非公式（数学区外）的可见文本区间强制显式写入 rPr（读后原值回写）。</summary>
        private void ForceExplicitRunProps(Paragraph paragraph)
        {
            try
            {
                Range whole = paragraph.Range;
                int mathCount;
                try { mathCount = whole.OMaths.Count; } catch { mathCount = 0; }
                if (mathCount == 0) return;

                // 数学区边界 → 非公式子区间
                var bounds = new List<KeyValuePair<int, int>>();
                for (int i = 1; i <= mathCount; i++)
                {
                    try
                    {
                        Range mr = whole.OMaths[i].Range;
                        bounds.Add(new KeyValuePair<int, int>(mr.Start, mr.End));
                    }
                    catch { }
                }
                bounds.Sort((a, b) => a.Key.CompareTo(b.Key));

                var gaps = new List<Range>();
                int start = whole.Start;
                foreach (var b in bounds)
                {
                    if (b.Key > start)
                    {
                        Range gap = whole.Duplicate;
                        gap.Start = start;
                        gap.End = b.Key;
                        if (HasVisible(gap)) gaps.Add(gap);
                    }
                    if (b.Value > start) start = b.Value;
                }
                if (whole.End > start)
                {
                    Range tail = whole.Duplicate;
                    tail.Start = start;
                    tail.End = whole.End;
                    if (HasVisible(tail)) gaps.Add(tail);
                }

                foreach (Range g in gaps)
                {
                    try
                    {
                        Font f = g.Font;
                        // 读后原值回写：强制 Word 序列化出显式 w:rFonts/w:sz/w:b/w:i/w:color
                        string fe = f.NameFarEast; if (!string.IsNullOrEmpty(fe)) f.NameFarEast = fe;
                        string la = f.NameAscii; if (!string.IsNullOrEmpty(la)) f.NameAscii = la;
                        string lo = f.NameOther; if (!string.IsNullOrEmpty(lo)) f.NameOther = lo;
                        float sz = f.Size;
                        if (sz > 0 && sz < WordFormatUtils.UndefinedFloat) f.Size = sz;
                        int bold = f.Bold;
                        if (bold == -1 || bold == 0) f.Bold = bold;
                        int ital = f.Italic;
                        if (ital == -1 || ital == 0) f.Italic = ital;
                        int col = (int)f.Color;
                        if (col >= 0 && col <= 0xFFFFFF) f.Color = (WdColor)col;
                    }
                    catch { }
                }
            }
            catch { }
        }

        private static bool HasVisible(Range r)
        {
            try
            {
                string t = r.Text ?? "";
                foreach (char ch in t)
                {
                    if (ch != '\r' && ch != '\a' && ch != '\v' && ch != ' ' && ch != '　') return true;
                }
            }
            catch { }
            return false;
        }

        private static int ParagraphIndex(Paragraph p)
        {
            // 段落序号用于失败定位（基于文档段落集合，代价可接受）
            try { return p.Range.Paragraphs[1].Range.Start; }
            catch { return -1; }
        }

        /// <summary>片段预备（批量注入用）：悬挂括号剥离/平衡校验/映射/MathML 构建。
        /// 与 TryConvertRange 的前置同语义（改动需两处同步）；不可注入返回 null。</summary>
        private string PrepareMml(string originalText)
        {
            string trimmed = TrimDanglingDelimiters(originalText);
            if (trimmed.Length == 0) return null;
            if (HasUnbalancedDelimiters(trimmed)) return null;
            string mapped = UnicodeMathMapper.NormalizeMathAlphanumerics(UnicodeMathMapper.Map(trimmed));
            try { return MathMLBuilder.Build(mapped); }
            catch (Exception bex) { RecordFallback("BUILD:" + mapped + " @" + bex.Message); return null; }
        }

        /// <summary>单个片段的转换。失败时尽力还原原文并返回 false。scratch：复用草稿（可空）。</summary>
        private bool TryConvertRange(Range rng, string originalText, bool secondPass = false, int segStart = -1, int segLen = 0, int paraLen = -1, Document scratch = null)
        {
            // 边缘悬挂括号剥离：尾部悬 ( [（后随中文说明被切开的场景）与头部悬 ) 剥掉再判平衡；
            // 剥离后仍不平衡（如 y = a(x - h）才判失败保留原文（TC-11）
            string trimmed = TrimDanglingDelimiters(originalText);
            if (trimmed.Length == 0) return false;
            if (HasUnbalancedDelimiters(trimmed)) return false;
            originalText = trimmed;

            string mapped = UnicodeMathMapper.NormalizeMathAlphanumerics(UnicodeMathMapper.Map(originalText));
            int mathCountBefore = 0;
            try { mathCountBefore = rng.OMaths.Count; } catch { }
            int docZonesBefore = 0;
            try { docZonesBefore = _doc.OMaths.Count; } catch { }

            try
            {
                // TC-19 管线（PRD §6.2.1）：MathML → 草稿文档 → FormattedText 注入（嵌套结构完整）
                // 二遍清扫不用草稿（漂移场景复发），直接走确定性 Functions/BuildUp 路径
                if (!secondPass && !string.IsNullOrEmpty(_mml2OmmlPath))
                {
                    string mml = null;
                    try { mml = MathMLBuilder.Build(mapped); }
                    catch (Exception bex) { RecordFallback("BUILD:" + mapped + " @" + bex.Message); }
                    if (mml != null)
                    {
                        if (MathmlInjector.TryInject(_doc, rng, mml, _mml2OmmlPath, segStart, segLen, paraLen, scratch)) return true;
                        RecordFallback("INJECT:" + mapped);
                    }
                }

                if (StructuredMath.NeedsStructure(mapped))
                    return ConvertStructured(rng, mapped, mathCountBefore, docZonesBefore);
                return ConvertPlain(rng, mapped, mathCountBefore);
            }
            catch
            {
                Restore(rng, originalText, mathCountBefore);
                return false;
            }
        }

        /// <summary>无结构片段：线性文本 → OMaths.Add → BuildUp（构建上下标）。</summary>
        private bool ConvertPlain(Range rng, string mapped, int mathCountBefore)
        {
            if (!string.IsNullOrEmpty(mapped)) rng.Text = mapped;

            Range mathRange = _doc.OMaths.Add(rng);
            OMath om = ZoneMath(mathRange);
            if (om == null) throw new InvalidOperationException("未取到公式对象。");
            om.BuildUp();

            int after = SafeMathCount(rng);
            if (after <= mathCountBefore) throw new InvalidOperationException("公式对象未建立。");
            ApplyMathSize(om);
            return true;
        }

        /// <summary>
        /// 结构化片段（顶层分式 / 根式）：BuildUp 只构建上下标、不建 m:f/m:rad（探针实证），
        /// 走 OMathFunctions.Add(Frac/Rad)：规范化分式括号 → 建区 → 根式 span 从右到左
        /// 逐个 Rad → 顶层 Frac → BuildUp 收尾剩余脚本。
        /// </summary>
        private bool ConvertStructured(Range rng, string mapped, int mathCountBefore, int docZonesBefore = 0)
        {
            string work = StructuredMath.NormalizeFractionParens(mapped);
            if (!string.IsNullOrEmpty(work)) rng.Text = work;

            int zoneStart = rng.Start;
            Range mathRange = _doc.OMaths.Add(rng);
            OMath om = ZoneMath(mathRange);
            if (om == null) throw new InvalidOperationException("未取到公式对象。");

            Action buildUp = () => { try { om.BuildUp(); } catch { } };
            Action rad = () =>
            {
                foreach (var span in StructuredMath.RadicalSpans(work))
                {
                    try
                    {
                        Range sub = _doc.Range(zoneStart + span.Key, zoneStart + span.Value);
                        om.Functions.Add(sub, WdOMathFunctionType.wdOMathFunctionRad);
                    }
                    catch { }
                }
            };
            Action frac = () =>
            {
                var fracSpan = StructuredMath.FractionSpan(work);
                if (fracSpan.HasValue)
                {
                    try
                    {
                        Range sub = _doc.Range(zoneStart + fracSpan.Value.Key, zoneStart + fracSpan.Value.Value);
                        om.Functions.Add(sub, WdOMathFunctionType.wdOMathFunctionFrac);
                    }
                    catch { }
                }
            };

            Action fracWhole = () =>
            {
                try { om.Functions.Add(om.Range, WdOMathFunctionType.wdOMathFunctionFrac); }
                catch { }
            };
            Action noBuild = () => { };

            // 最终策略（矩阵实证）：
            // 有顶层斜杠 → Rad(生文本)→Frac(整区,循环至无斜杠)→BuildUp（分式必须吃生文本，否则产物非法不可保存）
            // 无斜杠 → BuildUp(脚本全建,含√组内)→Rad（对已建内容包根式,验证可保存）
            int expectedFracs = StructuredMath.CountTopLevelSlashes(work);
            Action fracLoop = () =>
            {
                for (int i = 0; i < expectedFracs; i++)
                {
                    try { om.Functions.Add(om.Range, WdOMathFunctionType.wdOMathFunctionFrac); }
                    catch { break; }
                }
            };
            bool hasSlash = StructuredMath.TopLevelSlashIndex(work) >= 0;
            switch (StructureOrderMode)
            {
                case 0: case 3:
                    if (hasSlash) { rad(); fracLoop(); buildUp(); }
                    else { buildUp(); rad(); }
                    break;
                case 1: buildUp(); rad(); fracWhole(); break;
                case 2: rad(); buildUp(); frac(); break;
                case 5: buildUp(); rad(); fracWhole(); break;
                default:
                    if (hasSlash) { rad(); fracLoop(); buildUp(); }
                    else { buildUp(); rad(); }
                    break;
            }

            int after = SafeMathCount(rng);
            if (after <= mathCountBefore) throw new InvalidOperationException("公式对象未建立。");
            ApplyMathSize(om);
            return true;
        }


        private static string ZoneTextSafe(OMath om)
        {
            try { return om.Range.Text ?? ""; }
            catch { return ""; }
        }

        private static OMath ZoneMath(Range mathRange)
        {
            try { return mathRange.OMaths[1]; }
            catch { return null; }
        }

        private static int SafeMathCount(Range rng)
        {
            try { return rng.OMaths.Count; }
            catch { return 1; }
        }

        private void ApplyMathSize(OMath om)
        {
            if (_mathSizePt.HasValue)
            {
                try { om.Range.Font.Size = (float)_mathSizePt.Value; } catch { }
            }
        }


        /// <summary>剥离片段边缘悬挂的开/闭括号（尾部未配对 ( [ { 与头部未配对 ) ] }）。</summary>
        private static string TrimDanglingDelimiters(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            string t = s.TrimEnd();
            while (t.Length > 0 && (t[t.Length - 1] == '(' || t[t.Length - 1] == '[' || t[t.Length - 1] == '{'))
                t = t.Substring(0, t.Length - 1).TrimEnd();
            t = t.TrimStart();
            while (t.Length > 0 && (t[0] == ')' || t[0] == ']' || t[0] == '}'))
                t = t.Substring(1).TrimStart();
            return t;
        }

        /// <summary>圆括号/方括号/花括号是否不闭合（不成对即视为不闭合）。</summary>
        private static bool HasUnbalancedDelimiters(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            int paren = 0, bracket = 0, brace = 0;
            foreach (char c in s)
            {
                switch (c)
                {
                    case '(': paren++; break;
                    case ')': paren--; break;
                    case '[': bracket++; break;
                    case ']': bracket--; break;
                    case '{': brace++; break;
                    case '}': brace--; break;
                }
            }
            // 仅未闭合的开括号判失败（TC-11 的 y = a(x - h）；
            // 游离闭括号（如 m) ×）由 Word 字面渲染，放行（黄金取证）
            return paren > 0 || bracket > 0 || brace > 0;
        }

        /// <summary>失败恢复：移除可能残留的公式对象并还原纯文本。</summary>
        private void Restore(Range rng, string originalText, int mathCountBefore)
        {
            try
            {
                int count;
                try { count = rng.OMaths.Count; } catch { count = 0; }
                for (int i = count; i > mathCountBefore; i--)
                {
                    try { rng.OMaths[i].Remove(); } catch { }
                }
                rng.Text = originalText;
            }
            catch
            {
                // 还原失败是极端情况：不让异常冒泡打断整体转换
            }
        }

        // ---- MML2OMML.XSL 兜底管线（PRD §6.2.1，P2 结构化生成器接入点）----
        private static XslCompiledTransform _mmlTransform;
        private static string _mmlTransformPath;

        public static string MathmlToOmml(string mathml, string xslPath)
        {
            if (string.IsNullOrEmpty(xslPath) || !File.Exists(xslPath))
                throw new FileNotFoundException("MML2OMML.XSL 未找到。", xslPath);
            if (_mmlTransform == null || _mmlTransformPath != xslPath)
            {
                var transform = new XslCompiledTransform();
                transform.Load(xslPath);
                _mmlTransform = transform;
                _mmlTransformPath = xslPath;
            }
            using (var reader = new StringReader(mathml))
            using (var writer = new StringWriter())
            {
                using (var xmlReader = System.Xml.XmlReader.Create(reader))
                {
                    _mmlTransform.Transform(xmlReader, new XsltArgumentList(), writer);
                }
                return writer.ToString();
            }
        }

        public static string WrapLinearAsMtextMathml(string linearText)
        {
            string escaped = (linearText ?? "")
                .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
            return "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mtext>" + escaped + "</mtext></math>";
        }
    }
}
