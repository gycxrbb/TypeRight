using System;
using System.Collections.Generic;
using System.Linq;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Mixing;
using TypeRight.Profiles;
using TypeRight.Roles;
using TypeRight.Scanning;

namespace TypeRight.Formatting
{
    /// <summary>一次排版请求。</summary>
    public sealed class FormatRequest
    {
        public Profile Profile;
        /// <summary>段落索引（文档序，1 起）→ 角色。来自预览界面，允许用户改判。</summary>
        public IDictionary<int, ParagraphRole> RoleAssignments;
        public bool MountStyles = true;
        public bool ConvertFormulas = true;
        /// <summary>非 null 时只处理选区内段落（选区模式）。</summary>
        public Range Selection;
    }

    /// <summary>
    /// 排版执行器（PRD §5）：UndoRecord 打包为单次 Ctrl+Z；失败 EndCustomRecord + 一次 Undo 整体回退；
    /// 幂等（SetIfChanged 读后写）；保留局部强调（只动字体族/字号与段落级属性，不碰区间内加粗/颜色）。
    /// </summary>
    public sealed class FormatApplier
    {
        private readonly StyleMounter _mounter = new StyleMounter();

        /// <summary>诊断开关（冒烟测试/调试用）。</summary>
        public static bool DebugLog;

        private static void Log(string message)
        {
            if (DebugLog) Console.WriteLine("    [applier] " + message);
        }

        public ApplyStats Apply(Application app, Document doc, FormatRequest request)
        {
            if (app == null) throw new ArgumentNullException("app");
            if (doc == null) throw new ArgumentNullException("doc");
            if (request == null || request.Profile == null) throw new ArgumentNullException("request");
            if (request.RoleAssignments == null || request.RoleAssignments.Count == 0)
                throw new ArgumentException("RoleAssignments 为空。");

            var stats = new ApplyStats();
            bool screenOff = false;
            try { app.ScreenUpdating = false; screenOff = true; } catch { }

            UndoRecord undo = app.UndoRecord;
            undo.StartCustomRecord("TypeRight 智能排版");
            try
            {
                ApplyCore(app, doc, request, stats);
            }
            catch (Exception ex)
            {
                // 不破坏原则：任何失败整体回退（自定义撤销记录是一个撤销单元）
                try { undo.EndCustomRecord(); } catch { }
                try { doc.Undo(); } catch { }
                throw new InvalidOperationException("排版执行失败，已整体回退：" + ex.Message, ex);
            }
            finally
            {
                try { undo.EndCustomRecord(); } catch { }
                if (screenOff) { try { app.ScreenUpdating = true; } catch { } }
            }
            return stats;
        }

        private void ApplyCore(Application app, Document doc, FormatRequest request, ApplyStats stats)
        {
            Paragraphs paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;

            // 按文档序收集段落对象与角色（与 DocumentScanner 的索引口径一致）
            var targets = new List<KeyValuePair<Paragraph, ParagraphRole>>();
            for (int i = 1; i <= count; i++)
            {
                Paragraph p = null;
                try { p = paragraphs[i]; } catch { }
                if (p == null) continue;

                ParagraphRole role;
                if (!request.RoleAssignments.TryGetValue(i, out role)) continue;

                if (role == ParagraphRole.Skip)
                {
                    stats.Skipped++;
                    continue;
                }
                stats.CountRole(role);

                if (request.Selection != null && !InRange(p, request.Selection)) continue;

                targets.Add(new KeyValuePair<Paragraph, ParagraphRole>(p, role));
            }

            var styleCache = new Dictionary<string, Style>();
            foreach (var group in targets.GroupBy(kv => kv.Value))
            {
                ParagraphRole role = group.Key;
                RoleFormat fmt = ResolveFormat(request.Profile, role);
                bool mounted = false;
                Style style = null;

                if (request.MountStyles && fmt != null && !string.IsNullOrEmpty(fmt.WordStyle))
                {
                    string cacheKey = fmt.WordStyle;
                    if (!styleCache.TryGetValue(cacheKey, out style))
                    {
                        style = _mounter.EnsureStyle(doc, fmt.WordStyle);
                        styleCache[cacheKey] = style;
                    }
                    if (style != null)
                    {
                        _mounter.WriteStyleDefinition(doc, style, fmt, stats);
                        mounted = true;
                    }
                }

                foreach (var kv in group)
                {
                    Paragraph p = kv.Key;
                    bool changed = false;

                    if (mounted)
                        changed |= _mounter.MountParagraph(p, style);
                    if (fmt != null)
                        changed |= ApplyDirectFormat(p, fmt, mounted);
                    if (changed)
                    {
                        stats.ChangedParagraphs++;
                        Log("#" + ((Range)p.Range).Start + " " + role + " 改动");
                    }
                }
            }

            // 混排规则（在公式转换前，作用于当前纯文本）
            var mixing = request.Profile.MixingRules ?? new MixingRuleOptions();
            if (mixing.AutoSpaceCjkLatin || mixing.PunctuationMode != PunctuationMode.None)
            {
                foreach (var kv in targets)
                {
                    if (kv.Value == ParagraphRole.FormulaLine) continue;
                    ApplyMixingRules(kv.Key, mixing);
                }
            }

            // 公式转换
            if (request.ConvertFormulas && request.Profile.Formula != null && request.Profile.Formula.ConvertPlainText)
            {
                string xslPath = FormulaConverter.ResolveMml2OmmlPath(app);
                var converter = new FormulaConverter(doc, request.Profile.Formula, xslPath);
                var convertibleRoles = new HashSet<ParagraphRole>
                {
                    ParagraphRole.Body, ParagraphRole.FormulaLine, ParagraphRole.ListItem,
                    ParagraphRole.Caption, ParagraphRole.TableText, ParagraphRole.TableHeader
                };
                foreach (var kv in targets)
                {
                    if (!convertibleRoles.Contains(kv.Value)) continue;
                    // 表格单元格内的公式转换在 v1 跳过（区间边界复杂度），P1 打开
                    bool wholeLine = kv.Value == ParagraphRole.FormulaLine;
                    try
                    {
                        string text = kv.Key.Range.Text ?? "";
                        var f = new ParagraphFeatures { Text = text.TrimEnd('\r', '\a', '\v') };
                        f.HasFields = false;
                        converter.ConvertParagraph(f, kv.Key, stats, wholeLine);
                    }
                    catch { /* 单段失败不影响整体 */ }
                }
            }
        }

        private static bool InRange(Paragraph p, Range selection)
        {
            try { return p.Range.Start >= selection.Start && p.Range.End <= selection.End; }
            catch { return false; }
        }

        /// <summary>角色 → RoleFormat。formulaLine 缺省回落 body；缺失返回 null（只挂公式不刷格式）。</summary>
        private static RoleFormat ResolveFormat(Profile profile, ParagraphRole role)
        {
            RoleFormat fmt;
            if (profile.Roles != null && profile.Roles.TryGetValue(role.FormatKey(), out fmt)) return fmt;
            if (role == ParagraphRole.FormulaLine && profile.Roles != null && profile.Roles.TryGetValue(ParagraphRole.Body.ToProfileKey(), out fmt))
                return fmt;
            return null;
        }

        /// <summary>
        /// 段落级直接格式写入（幂等）。
        /// 字体族/字号只作用于"数学区之外的文本子区间"（公式对象由转换器负责，避免读写数学区字体破坏幂等）；
        /// 挂样式时优先把直接字体格式"清除回退到样式"（wdUndefined 赋值，首次运行自探测；失败则直写目标值兜底）。
        /// 段落级对齐/行距/缩进/间距一律直接写。区间内的局部加粗/颜色完全不触碰（保留强调）。
        /// </summary>
        private bool ApplyDirectFormat(Paragraph p, RoleFormat fmt, bool mounted)
        {
            bool changed = false;
            foreach (Range textRange in TextRangesExcludingMath(p))
            {
                Font font = null;
                try { font = textRange.Font; } catch { }
                if (font == null) continue;
                if (mounted)
                    changed |= ClearDirectFonts(font, fmt);
                else
                    changed |= SetDirectFonts(font, fmt);
            }

            // 段落格式（直接写，覆盖既有直接格式）
            try
            {
                ParagraphFormat pf = p.Format;
                if (fmt.AlignmentValue != ProfileAlignment.None)
                    changed |= WordFormatUtils.SetIfChanged(() => p.Alignment, v => p.Alignment = v,
                        WordFormatUtils.ToWdAlignment(fmt.AlignmentValue), (a, b) => a == b);
                if (fmt.LineSpacing.HasValue)
                    WordFormatUtils.ApplyLineSpacing(pf, fmt.LineSpacing.Value);
                if (fmt.SpaceBeforePt.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => pf.SpaceBefore, v => pf.SpaceBefore = v, (float)fmt.SpaceBeforePt.Value, WordFormatUtils.FloatEq);
                if (fmt.SpaceAfterPt.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => pf.SpaceAfter, v => pf.SpaceAfter = v, (float)fmt.SpaceAfterPt.Value, WordFormatUtils.FloatEq);
                if (fmt.FirstLineIndentPt.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => pf.FirstLineIndent, v => pf.FirstLineIndent = v, (float)fmt.FirstLineIndentPt.Value, WordFormatUtils.FloatEq);
                if (fmt.OutlineLevel.HasValue && (fmt.WordStyle == null || !mounted))
                    changed |= WordFormatUtils.SetIfChanged(() => p.OutlineLevel, v => p.OutlineLevel = v,
                        WordFormatUtils.ToWdOutlineLevel(fmt.OutlineLevel.Value), (a, b) => a == b);
            }
            catch { }

            return changed;
        }

        /// <summary>段落内数学区之外的文本子区间（含段落标记）。无公式时返回整个段落区间。</summary>
        private static List<Range> TextRangesExcludingMath(Paragraph p)
        {
            var result = new List<Range>();
            Range whole = p.Range;
            int mathCount = 0;
            try { mathCount = whole.OMaths.Count; } catch { }
            if (mathCount == 0)
            {
                result.Add(whole);
                return result;
            }

            // 收集数学区边界（按起点排序）
            var bounds = new List<KeyValuePair<int, int>>();
            for (int i = 1; i <= mathCount; i++)
            {
                try
                {
                    var mr = whole.OMaths[i].Range;
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
                    var gap = whole.Duplicate;
                    gap.Start = start;
                    gap.End = b.Key;
                    result.Add(gap);
                }
                if (b.Value > start) start = b.Value;
            }
            if (whole.End > start)
            {
                var tail = whole.Duplicate;
                tail.Start = start;
                tail.End = whole.End;
                result.Add(tail);
            }
            return result;
        }

        // ---- 直接格式清除：wdUndefined 赋值自探测 ----
        // null=未探测；true=undefined 赋值可清除直接格式；false=不可（兜底直写目标值）

        private static bool? _undefinedSizeClearWorks;
        private static bool? _undefinedNameClearWorks;

        private bool ClearDirectFonts(Font font, RoleFormat fmt)
        {
            bool changed = false;

            // ---- 字号 ----
            float sizeBefore = SafeRead(() => font.Size);
            if (fmt.SizePt.HasValue)
            {
                if (_undefinedSizeClearWorks == null)
                {
                    try
                    {
                        font.Size = WordFormatUtils.UndefinedFloat;
                        float probe = font.Size;
                        _undefinedSizeClearWorks = Math.Abs(probe - (float)fmt.SizePt.Value) < 0.05f;
                    }
                    catch
                    {
                        _undefinedSizeClearWorks = false;
                    }
                }

                if (_undefinedSizeClearWorks == true)
                {
                    try { font.Size = WordFormatUtils.UndefinedFloat; }
                    catch
                    {
                        WordFormatUtils.SetIfChanged(() => font.Size, v => font.Size = v, (float)fmt.SizePt.Value, WordFormatUtils.FloatEq);
                    }
                }
                else
                {
                    WordFormatUtils.SetIfChanged(() => font.Size, v => font.Size = v, (float)fmt.SizePt.Value, WordFormatUtils.FloatEq);
                }
            }
            float sizeAfter = SafeRead(() => font.Size);
            changed |= Math.Abs(sizeAfter - sizeBefore) > 0.05f && !(sizeBefore >= WordFormatUtils.UndefinedFloat && sizeAfter >= WordFormatUtils.UndefinedFloat);

            // ---- 字体族 ----
            string farEastBefore = SafeRead(() => font.NameFarEast);
            string asciiBefore = SafeRead(() => font.NameAscii);
            bool hasNames = !string.IsNullOrEmpty(fmt.EastAsiaFont) || !string.IsNullOrEmpty(fmt.LatinFont);
            if (hasNames)
            {
                if (_undefinedNameClearWorks == null)
                {
                    try
                    {
                        font.NameAscii = "";
                        string probe = font.NameAscii;
                        _undefinedNameClearWorks = probe != null &&
                            string.Equals(probe, fmt.LatinFont, StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        _undefinedNameClearWorks = false;
                    }
                }

                if (_undefinedNameClearWorks == true)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(fmt.EastAsiaFont)) font.NameFarEast = "";
                        if (!string.IsNullOrEmpty(fmt.LatinFont)) { font.NameAscii = ""; font.NameOther = ""; }
                    }
                    catch
                    {
                        SetDirectFonts(font, fmt);
                    }
                }
                else
                {
                    SetDirectFonts(font, fmt);
                }
            }
            string farEastAfter = SafeRead(() => font.NameFarEast);
            string asciiAfter = SafeRead(() => font.NameAscii);
            changed |= !string.Equals(farEastBefore, farEastAfter, StringComparison.Ordinal) ||
                       !string.Equals(asciiBefore, asciiAfter, StringComparison.Ordinal);

            return changed;
        }

        private static T SafeRead<T>(Func<T> read)
        {
            try { return read(); }
            catch { return default(T); }
        }

        private bool SetDirectFonts(Font font, RoleFormat fmt)
        {
            bool changed = false;
            if (!string.IsNullOrEmpty(fmt.EastAsiaFont))
                changed |= WordFormatUtils.SetIfChanged(() => font.NameFarEast, v => font.NameFarEast = v, fmt.EastAsiaFont);
            if (!string.IsNullOrEmpty(fmt.LatinFont))
            {
                changed |= WordFormatUtils.SetIfChanged(() => font.NameAscii, v => font.NameAscii = v, fmt.LatinFont);
                changed |= WordFormatUtils.SetIfChanged(() => font.NameOther, v => font.NameOther = v, fmt.LatinFont);
            }
            return changed;
        }

        /// <summary>混排规则：对段落文本计算编辑并降序写入 Range。</summary>
        private static void ApplyMixingRules(Paragraph p, MixingRuleOptions mixing)
        {
            try
            {
                Range range = p.Range;
                string text = range.Text;
                if (string.IsNullOrEmpty(text)) return;
                var edits = AutoSpacer.ComputeEdits(text, mixing.AutoSpaceCjkLatin, mixing.PunctuationMode);
                for (int i = edits.Count - 1; i >= 0; i--)
                {
                    var e = edits[i];
                    if (e.Position >= text.Length) continue;
                    Range sub = range.Duplicate;
                    sub.Start = range.Start + e.Position;
                    sub.End = sub.Start + e.Length;
                    sub.Text = e.Insert;
                }
            }
            catch { /* 混排失败不阻塞排版 */ }
        }
    }
}
