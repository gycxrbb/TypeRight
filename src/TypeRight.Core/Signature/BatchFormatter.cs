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

                if (c.EastAsiaFont != null)
                    changed |= WordFormatUtils.SetIfChanged(() => f.NameFarEast, v => f.NameFarEast = v, c.EastAsiaFont);
                if (c.LatinFont != null)
                {
                    changed |= WordFormatUtils.SetIfChanged(() => f.NameAscii, v => f.NameAscii = v, c.LatinFont);
                    changed |= WordFormatUtils.SetIfChanged(() => f.NameOther, v => f.NameOther = v, c.LatinFont);
                }
                if (c.SizePt.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => f.Size, v => f.Size = v, (float)c.SizePt.Value, WordFormatUtils.FloatEq);
                if (c.Bold.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => f.Bold, v => f.Bold = v, c.Bold.Value ? -1 : 0, WordFormatUtils.IntEq);
                if (c.ColorRgb.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => f.Color, v => f.Color = v,
                        (WdColor)WordFormatUtils.RgbToWdColor(c.ColorRgb.Value), (a, b) => a == b);
            }
            return changed;
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
                    result.Add(gap);
                }
                if (b.Value > start) start = b.Value;
            }
            if (whole.End > start)
            {
                Range tail = whole.Duplicate;
                tail.Start = start;
                tail.End = whole.End;
                result.Add(tail);
            }
            return result;
        }
    }
}
