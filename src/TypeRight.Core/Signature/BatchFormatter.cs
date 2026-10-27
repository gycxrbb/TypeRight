using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formatting;

namespace TypeRight.Signature
{
    /// <summary>一次批量修改的属性集。null = 保持源值不触碰（PRD v2 §8 apply 语义）。</summary>
    public sealed class FormatChanges
    {
        public string EastAsiaFont;
        /// <summary>西文联动写入 NameAscii/NameOther（P0-3 规范兜底）；null = 不联动。</summary>
        public string LatinFont;
        public double? SizePt;
        public bool? Bold;
        public int? ColorRgb;

        public bool HasAny
        {
            get
            {
                return EastAsiaFont != null || LatinFont != null || SizePt.HasValue || Bold.HasValue || ColorRgb.HasValue;
            }
        }
    }

    /// <summary>批量执行统计（幂等口径：值未变化的段落不计入 Changed）。</summary>
    public sealed class BatchStats
    {
        public int Total;
        public int Changed;
        public string UndoName;
    }

    /// <summary>
    /// 批量格式执行器（PRD v2 P0-2/P0-5）：每次调用 = 一个撤销单元；
    /// 只写显式给出的属性；含 OMath 区间的段落只在数学区外子区间写字体（数学区由公式引擎管理）。
    /// </summary>
    public static class BatchFormatter
    {
        /// <summary>幂等计数探针（验收调试用）。</summary>
        public static bool DiagCount;
        internal static int _diagPrinted;

        public static BatchStats Apply(Application app, Document doc, IList<Paragraph> paragraphs, FormatChanges changes,
            string undoName, bool wrapUndo = true)
        {
            if (changes == null || !changes.HasAny)
                return new BatchStats { Total = paragraphs == null ? 0 : paragraphs.Count, Changed = 0 };

            var stats = new BatchStats { Total = paragraphs == null ? 0 : paragraphs.Count, UndoName = undoName ?? "TypeRight 批量格式" };

            bool screenOff = false;
            try { app.ScreenUpdating = false; screenOff = true; } catch { }

            UndoRecord undo = null;
            if (wrapUndo)
            {
                undo = app.UndoRecord;
                undo.StartCustomRecord(stats.UndoName);
            }
            try
            {
                foreach (Paragraph p in paragraphs)
                {
                    try
                    {
                        if (ApplyToParagraph(doc, p, changes)) stats.Changed++;
                    }
                    catch { /* 单段失败跳过，不中断批次 */ }
                }
            }
            finally
            {
                if (undo != null) { try { undo.EndCustomRecord(); } catch { } }
                if (screenOff) { try { app.ScreenUpdating = true; } catch { } }
            }
            return stats;
        }

        /// <summary>计算改动但不执行（P2 排版体检预留：与执行共用同一判定逻辑）。</summary>
        public static int CountPendingChanges(Document doc, IList<Paragraph> paragraphs, FormatChanges changes)
        {
            if (changes == null || !changes.HasAny || paragraphs == null) return 0;
            int pending = 0;
            foreach (Paragraph p in paragraphs)
            {
                try
                {
                    if (WouldChange(p, changes)) pending++;
                }
                catch { }
            }
            return pending;
        }

        private static bool WouldChange(Paragraph p, FormatChanges c)
        {
            foreach (Range r in TextRangesExcludingMath(p))
            {
                Font f = null;
                try { f = r.Font; } catch { }
                if (f == null) continue;
                if (c.EastAsiaFont != null && !string.Equals(f.NameFarEast, c.EastAsiaFont, StringComparison.OrdinalIgnoreCase)) return true;
                if (c.LatinFont != null && !string.Equals(f.NameAscii, c.LatinFont, StringComparison.OrdinalIgnoreCase)) return true;
                if (c.SizePt.HasValue && Math.Abs(f.Size - (float)c.SizePt.Value) > 0.05f) return true;
                if (c.Bold.HasValue && f.Bold != (c.Bold.Value ? -1 : 0)) return true;
                if (c.ColorRgb.HasValue && WordFormatUtils.WdColorToRgb((int)f.Color) != c.ColorRgb.Value) return true;
            }
            return false;
        }

        private static bool ApplyToParagraph(Document doc, Paragraph p, FormatChanges c)
        {
            bool changed = false;
            foreach (Range r in TextRangesExcludingMath(p))
            {
                Font f = null;
                try { f = r.Font; } catch { }
                if (f == null) continue;

                // 前置快照（幂等计数与字号护栏用）
                string beforeFE = SafeFontName(f.NameFarEast);
                string beforeAscii = SafeFontName(f.NameAscii);
                float beforeSize = SafeSize(f.Size);

                // 字号护栏：区间级字体族写入会重写 rPr，个别场景剥掉 w:sz（★ 段实测）。
                // 有效字号先记下，写完字体后缺失则回写。
                bool guardSize = beforeSize > 0 && beforeSize < WordFormatUtils.UndefinedFloat;

                if (c.EastAsiaFont != null)
                    WordFormatUtils.SetIfChanged(() => f.NameFarEast, v => f.NameFarEast = v, c.EastAsiaFont);
                if (c.LatinFont != null)
                {
                    WordFormatUtils.SetIfChanged(() => f.NameAscii, v => f.NameAscii = v, c.LatinFont);
                    WordFormatUtils.SetIfChanged(() => f.NameOther, v => f.NameOther = v, c.LatinFont);
                }
                if (c.SizePt.HasValue)
                    WordFormatUtils.SetIfChanged(() => f.Size, v => f.Size = v, (float)c.SizePt.Value, WordFormatUtils.FloatEq);
                if (c.Bold.HasValue)
                    WordFormatUtils.SetIfChanged(() => f.Bold, v => f.Bold = v, c.Bold.Value ? -1 : 0, WordFormatUtils.IntEq);
                if (c.ColorRgb.HasValue)
                    WordFormatUtils.SetIfChanged(() => f.Color, v => f.Color = v,
                        (WdColor)WordFormatUtils.RgbToWdColor(c.ColorRgb.Value), (a, b) => a == b);

                if (guardSize)
                {
                    // 无条件回写：rPr 在序列化时可能丢弃 w:sz（API 读不出），已知字号一律显式回写
                    try { f.Size = beforeSize; } catch { }
                }

                // 幂等计数：before 为已知一致值且 after 不同才计改动；
                // before 读数混合（含段落/单元格标记）→ 归一写不算改动（幂等口径）
                string afterFE = SafeFontName(f.NameFarEast);
                string afterAscii = SafeFontName(f.NameAscii);
                float afterSize2 = SafeSize(f.Size);
                // 只对本次写入过的维度计改动；读数为混合（null，写后同代理常见陈旧值）一律不计
                bool feDiff = c.EastAsiaFont != null && beforeFE != null && afterFE != null && !string.Equals(beforeFE, afterFE, StringComparison.OrdinalIgnoreCase);
                                bool asciiDiff = c.LatinFont != null && beforeAscii != null && afterAscii != null
                                        && !string.Equals(beforeAscii, afterAscii, StringComparison.OrdinalIgnoreCase)
                                        && !string.Equals(beforeAscii, "Cambria Math", StringComparison.OrdinalIgnoreCase)
                                        && !string.Equals(afterAscii, "Cambria Math", StringComparison.OrdinalIgnoreCase);
                bool sizeDiff = c.SizePt.HasValue && beforeSize > 0 && afterSize2 > 0 && Math.Abs(beforeSize - afterSize2) > 0.05f;
                if (feDiff || asciiDiff || sizeDiff)
                {
                    changed = true;
                    if (DiagCount && _diagPrinted < 8)
                    {
                        _diagPrinted++;
                        System.Console.WriteLine("    [idem-probe] FE[" + beforeFE + "]->[" + afterFE + "] Ascii[" + beforeAscii + "]->[" + afterAscii + "] Sz[" + beforeSize + "]->[" + afterSize2 + "]");
                    }
                }
            }
            return changed;
        }

        private static string SafeFontName(string name)
        {
            // null = Word 对混合区间的读数（不计入幂等改动）
            return string.IsNullOrWhiteSpace(name) ? null : name;
        }

        private static float SafeSize(float size)
        {
            return size >= WordFormatUtils.UndefinedFloat ? -1f : size;
        }


        /// <summary>区间是否含可见文本（纯段落/单元格标记的区间不参与字体写入——纯公式段的标记保持数学字体邻接态，避免幂等抖动）。</summary>
        private static bool HasVisibleText(Range r)
        {
            try
            {
                string t = r.Text ?? "";
                foreach (char ch in t)
                {
                    if (ch != (char)13 && ch != (char)7 && ch != (char)11 && ch != ' ' && ch != (char)0x3000) return true;
                }
            }
            catch { }
            return false;
        }

        /// <summary>段落内数学区之外的文本子区间（含段落标记）。无公式时返回整个段落。</summary>
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
                    if (HasVisibleText(gap)) result.Add(gap);
                }
                if (b.Value > start) start = b.Value;
            }
            if (whole.End > start)
            {
                Range tail = whole.Duplicate;
                tail.Start = start;
                tail.End = whole.End;
                if (HasVisibleText(tail)) result.Add(tail);
            }
            return result;
        }
    }
}
