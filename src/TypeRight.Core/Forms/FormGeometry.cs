using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Forms
{
    /// <summary>
    /// 表格与单元格几何测绘工具（零 DOM 写入纯数学反推坐标 + 最少 Information 调用）。
    /// </summary>
    public static class FormGeometry
    {
        public struct RectPt
        {
            public float Left;
            public float Top;
            public float Width;
            public float Height;
            public int Page;

            public float Right { get { return Left + Width; } }
            public float Bottom { get { return Top + Height; } }
            public bool IsValid { get { return Width > 2f && Height > 2f && Left >= 0f && Top >= 0f; } }
        }

        /// <summary>确保当前窗口处于页面视图（PrintView），否则 wdVerticalPositionRelativeToPage 返回 -1。</summary>
        public static void EnsurePrintView(Application app)
        {
            try
            {
                Window win = app.ActiveWindow;
                if (win != null && win.View != null && win.View.Type != WdViewType.wdPrintView)
                    win.View.Type = WdViewType.wdPrintView;
            }
            catch { }
        }

        /// <summary>单元格集合兼容枚举。</summary>
        public static List<Cell> CellsOf(Cells cells)
        {
            var list = new List<Cell>();
            if (cells == null) return list;
            try
            {
                foreach (object item in (System.Collections.IEnumerable)cells.UnderlyingObject)
                    list.Add(new Cell(null, item));
            }
            catch { }
            return list;
        }

        /// <summary>表格行集合兼容枚举。</summary>
        public static List<Row> RowsOf(Rows rows)
        {
            var list = new List<Row>();
            if (rows == null) return list;
            try
            {
                foreach (object item in (System.Collections.IEnumerable)rows.UnderlyingObject)
                    list.Add(new Row(null, item));
            }
            catch { }
            return list;
        }

        /// <summary>浮动图形集合兼容枚举。</summary>
        public static List<Shape> ShapesOf(Shapes shapes)
        {
            var list = new List<Shape>();
            if (shapes == null) return list;
            try
            {
                foreach (object item in (System.Collections.IEnumerable)shapes.UnderlyingObject)
                    list.Add(new Shape(null, item));
            }
            catch { }
            return list;
        }

        /// <summary>嵌入式图形集合兼容枚举。</summary>
        public static List<InlineShape> InlineShapesOf(InlineShapes inlineShapes)
        {
            var list = new List<InlineShape>();
            if (inlineShapes == null) return list;
            try
            {
                foreach (object item in (System.Collections.IEnumerable)inlineShapes.UnderlyingObject)
                    list.Add(new InlineShape(null, item));
            }
            catch { }
            return list;
        }

        /// <summary>清理 Word 单元格文本末尾自带的 "\r\a" 结束标记。</summary>
        public static string CleanCellText(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            return raw.TrimEnd('\r', '\a', '\n', ' ', '\t', '\u3000');
        }

        /// <summary>读取 Range 相对页面左上角的水平坐标（磅），失败返回 -1。</summary>
        public static float GetPageX(Range rng)
        {
            try
            {
                object val = rng.Information(WdInformation.wdHorizontalPositionRelativeToPage);
                return Convert.ToSingle(val);
            }
            catch { return -1f; }
        }

        /// <summary>读取 Range 相对页面左上角的垂直坐标（磅），失败返回 -1。</summary>
        public static float GetPageY(Range rng)
        {
            try
            {
                object val = rng.Information(WdInformation.wdVerticalPositionRelativeToPage);
                return Convert.ToSingle(val);
            }
            catch { return -1f; }
        }

        /// <summary>读取 Range 所在页码（1-based），失败返回 1。</summary>
        public static int GetPageNumber(Range rng)
        {
            try
            {
                object val = rng.Information(WdInformation.wdActiveEndPageNumber);
                return Convert.ToInt32(val);
            }
            catch { return 1; }
        }

        /// <summary>判断 Shape 是否启用了 LayoutInCell（锚定在单元格内时为 -1，锚定在表外段落时为 0 或忽略）。</summary>
        public static bool IsShapeLayoutInCell(Shape shape)
        {
            if (shape == null) return false;
            try { return shape.LayoutInCell != 0; }
            catch { return true; }
        }

        /// <summary>
        /// 获取位于表格外部且最靠近目标纵向位置的安全页面级锚点 Range（O(1)，零 Information 调用）。
        /// - 上半页单元格优先锚定在表头首段（doc.Range(0,0)），下半页单元格优先锚定在表尾段（table.Range.End），
        ///   确保锚点始终处于当前屏幕视口内，防止 tr.Select() 触发 Word 视口上下滚动抖动；
        /// - 锚定在表格外部可彻底避免触发底端对齐单元格（wdCellAlignVerticalBottom）的内容高度重算
        ///   （防止「负责人签字/年月日」上移），并消除单元格左侧灰色锚定条。
        /// </summary>
        public static Range GetPageAnchorOutsideTable(Document doc, Table table, Cell fallbackCell, int targetPage, float targetTopY, out bool anchoredInCell)
        {
            anchoredInCell = false;
            if (doc == null) return null;

            try
            {
                if (table != null && targetPage <= 1)
                {
                    int beforePos = table.Range.Start;
                    int afterPos = table.Range.End;
                    int docEnd = doc.Content.End;

                    // 下半页单元格（Top >= 320pt）优先挂在表格紧邻下方的表尾段落，避免视口在页面下半部时向上滚回页首
                    if (targetTopY >= 320f && afterPos < docEnd)
                    {
                        return doc.Range(afterPos, afterPos);
                    }
                    if (beforePos > 0)
                    {
                        return doc.Range(0, 0);
                    }
                    if (afterPos < docEnd)
                    {
                        return doc.Range(afterPos, afterPos);
                    }
                }
                else if (table != null)
                {
                    int afterPos = table.Range.End;
                    if (afterPos < doc.Content.End)
                    {
                        return doc.Range(afterPos, afterPos);
                    }
                }
            }
            catch { }

            if (fallbackCell != null)
            {
                anchoredInCell = true;
                try { return doc.Range(fallbackCell.Range.Start, fallbackCell.Range.Start); }
                catch { }
            }
            return doc.Range(0, 0);
        }

        /// <summary>
        /// 极速测定照片目标单元格宽高（已清空提示字并归零对齐后调用）：
        /// - 非纵向合并格：直接读 Width / Height（2 次属性读取，0 次 Information）；
        /// - 纵向合并格（如跨 1~3 行的照片格）：直接从 targetCol 位置向后扫描首个宽格/同列格，
        ///   用 2 次 GetPageY 测定真实渲染跨行高度（即使第 3 行被多行文字撑高也 100% 严丝合缝，彻底消除底部留白！）。
        /// </summary>
        public static void MeasurePhotoCellSizeAfterNormalize(Table table, Cell targetCell, out float width, out float height)
        {
            width = 0f;
            height = 0f;
            if (targetCell == null) return;

            try
            {
                float w = targetCell.Width;
                if (w > 2f && w < 2000f) width = w;
            }
            catch { }

            try
            {
                float ch = targetCell.Height;
                if (ch > 4f && ch < 10000f)
                {
                    height = ch;
                    return;
                }
            }
            catch { }

            height = MeasureMergedPhotoCellHeightFast(table, targetCell, width);
        }

        private static float MeasureMergedPhotoCellHeightFast(Table table, Cell targetCell, float targetW)
        {
            if (table == null || targetCell == null) return 110f;

            try
            {
                Document doc = targetCell.Range.Document;
                int tStart = targetCell.Range.Start;
                float startTopY = GetPageY(doc.Range(tStart, tStart));

                int targetCol = 1;
                try { targetCol = targetCell.ColumnIndex; } catch { }

                List<Cell> allCells = CellsOf(table.Range.Cells);
                int startScanIdx = Math.Max(0, targetCol);
                Cell belowCell = null;

                for (int i = startScanIdx; i < allCells.Count; i++)
                {
                    Cell c = allCells[i];
                    try
                    {
                        if (c.Range.Start <= tStart) continue;
                        float cw = c.Width;
                        // 下方行的宽格（如第 4 行「团支部意见」内容格，宽 378pt > 照片格宽 71pt * 1.5）或同列格即为合并结束行
                        if ((targetW > 10f && cw > targetW * 1.4f) || c.ColumnIndex >= targetCol)
                        {
                            belowCell = c;
                            break;
                        }
                    }
                    catch { }
                }

                if (belowCell != null && startTopY >= 0f)
                {
                    int bStart = belowCell.Range.Start;
                    Range bProbe = doc.Range(bStart, bStart);
                    float belowTopY = GetPageY(bProbe);
                    if (belowTopY > startTopY + 15f)
                    {
                        // 扣除下方单元格首行因固定行距（如 20pt 行距）或段前距产生的额外下移偏移（约 2.85pt），确保底边 100% 严丝合缝贴合底边框
                        float extraLeading = 0f;
                        try
                        {
                            ParagraphFormat pf = bProbe.ParagraphFormat;
                            float sb = pf.SpaceBefore;
                            if (sb > 0f && sb < 100f) extraLeading += sb;
                            WdLineSpacing rule = pf.LineSpacingRule;
                            if (rule == WdLineSpacing.wdLineSpaceExactly || rule == WdLineSpacing.wdLineSpaceAtLeast)
                            {
                                float ls = pf.LineSpacing;
                                if (ls > 14.3f && ls < 60f)
                                    extraLeading += (ls - 14.3f) * 0.5f;
                            }
                            else if (rule == WdLineSpacing.wdLineSpace1pt5)
                            {
                                extraLeading += 3.0f;
                            }
                        }
                        catch { }

                        return Math.Max(20f, (belowTopY - extraLeading) - startTopY);
                    }
                }
            }
            catch { }

            return 110.4f;
        }

        /// <summary>
        /// 高性能测定单元格页面几何矩形（专用于插入悬浮填表框，零 DOM 写入、仅 2 次 Information 调用）。
        /// </summary>
        public static RectPt MeasureCellRect(Table table, Cell targetCell, bool normalizeInPlace = false)
        {
            var rect = new RectPt { Left = -1f, Top = -1f, Width = 0f, Height = 0f, Page = 1 };
            if (targetCell == null) return rect;

            const float leftPad = 5.4f;
            const float topPad = 0f;
            const float bottomPad = 0f;

            try
            {
                float w = targetCell.Width;
                if (w > 2f && w < 2000f) rect.Width = w;
            }
            catch { }

            float nominalHeight = -1f;
            try
            {
                float ch = targetCell.Height;
                if (ch > 4f && ch < 10000f)
                    nominalHeight = ch;
            }
            catch { }

            MeasureCellTopLeftFast(
                targetCell, rect.Width, nominalHeight, leftPad, topPad, bottomPad,
                normalizeInPlace, out rect.Left, out rect.Top, out rect.Page);

            // 非纵向合并单元格（nominalHeight > 4f）：无需枚举全表！
            if (nominalHeight > 4f)
            {
                if (nominalHeight >= 45f)
                {
                    rect.Height = nominalHeight;
                    return rect;
                }
                try
                {
                    int charLen = targetCell.Range.End - targetCell.Range.Start;
                    if (charLen <= 25)
                    {
                        rect.Height = nominalHeight;
                        return rect;
                    }
                    Document doc = targetCell.Range.Document;
                    int cEnd = Math.Max(targetCell.Range.Start, targetCell.Range.End - 1);
                    float yEnd = GetPageY(doc.Range(cEnd, cEnd));
                    if (yEnd >= rect.Top)
                    {
                        rect.Height = Math.Max(nominalHeight, (yEnd - rect.Top) + 15.6f + bottomPad);
                        return rect;
                    }
                }
                catch { }

                rect.Height = nominalHeight;
                return rect;
            }

            rect.Height = MeasureMergedPhotoCellHeightFast(table, targetCell, rect.Width);
            return rect;
        }

        private static void MeasureCellTopLeftFast(
            Cell cell,
            float cellWidth,
            float nominalHeight,
            float leftPad,
            float topPad,
            float bottomPad,
            bool normalizeInPlace,
            out float cellLeft,
            out float cellTop,
            out int cellPage)
        {
            cellLeft = -1f;
            cellTop = -1f;
            cellPage = 1;
            if (cell == null) return;

            Document doc = null;
            try { doc = cell.Range.Document; } catch { }

            WdCellVerticalAlignment vertAlign = WdCellVerticalAlignment.wdCellAlignVerticalTop;
            try { vertAlign = cell.VerticalAlignment; } catch { }

            if (normalizeInPlace)
            {
                try
                {
                    if (vertAlign != WdCellVerticalAlignment.wdCellAlignVerticalTop)
                    {
                        cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                        vertAlign = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    }
                    Paragraph p1 = cell.Range.Paragraphs[1];
                    if (p1.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                        p1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    if (Math.Abs(p1.LeftIndent) > 0.1f) p1.LeftIndent = 0f;
                    if (Math.Abs(p1.FirstLineIndent) > 0.1f) p1.FirstLineIndent = 0f;
                    if (Math.Abs(p1.SpaceBefore) > 0.1f) p1.SpaceBefore = 0f;
                }
                catch { }
            }

            try
            {
                int cStart = cell.Range.Start;
                Range startRng = doc != null ? doc.Range(cStart, cStart) : cell.Range.Duplicate;

                // 1. 读取水平坐标（仅 1 次 Information 调用）
                float xStart = GetPageX(startRng);
                float indentOffset = 0f;
                if (!normalizeInPlace && vertAlign == WdCellVerticalAlignment.wdCellAlignVerticalTop)
                {
                    try
                    {
                        indentOffset = Math.Max(0f, startRng.ParagraphFormat.FirstLineIndent);
                    }
                    catch { }
                }
                if (cellLeft < 0f)
                {
                    cellLeft = xStart >= 0f ? Math.Max(0f, xStart - leftPad - indentOffset) : -1f;
                }

                // 2. 读取垂直坐标（按垂直对齐方式仅执行 1 次 Information 调用！）
                if (vertAlign == WdCellVerticalAlignment.wdCellAlignVerticalBottom && nominalHeight > 4f)
                {
                    // 底端对齐单元格（如「学院团委意见」）：直接读末尾字符 Y 减去单元格高度，省去 startRng 的 Y 查询！
                    int cEnd = Math.Max(cStart, cell.Range.End - 1);
                    Range endRng = doc != null ? doc.Range(cEnd, cEnd) : cell.Range.Duplicate;
                    float yEnd = GetPageY(endRng);
                    if (yEnd >= 0f)
                        cellTop = Math.Max(0f, (yEnd + 15.6f + bottomPad) - nominalHeight);
                }
                else if (vertAlign == WdCellVerticalAlignment.wdCellAlignVerticalCenter && nominalHeight > 4f && cStart >= cell.Range.End - 2)
                {
                    // 单行垂直居中单元格：yStart 位于行中点
                    float yStart = GetPageY(startRng);
                    if (yStart >= 0f)
                        cellTop = Math.Max(0f, yStart - Math.Max(0f, (nominalHeight - 15.6f) * 0.5f) - topPad);
                }
                else
                {
                    float yStart = GetPageY(startRng);
                    if (yStart >= 0f)
                        cellTop = Math.Max(0f, yStart - topPad);
                }
            }
            catch { }
        }

        private static float MeasureCellTopYFast(Cell cell, float topPad, float bottomPad)
        {
            float w = 0f, h = -1f;
            try { w = cell.Width; } catch { }
            try { float ch = cell.Height; if (ch > 4f && ch < 10000f) h = ch; } catch { }
            float left, top;
            int page;
            MeasureCellTopLeftFast(cell, w, h, 5.4f, topPad, bottomPad, false, out left, out top, out page);
            return top;
        }

        private static float MeasureCellHeight(
            Table table,
            Cell targetCell,
            float cellLeft,
            float cellTopY,
            float cellWidth,
            int cellPage,
            float nominalCellHeight,
            float leftPad,
            float topPad,
            float bottomPad)
        {
            int targetRow = 1, targetCol = 1;
            try { targetRow = targetCell.RowIndex; targetCol = targetCell.ColumnIndex; } catch { }

            float cellRight = cellLeft >= 0f && cellWidth > 0f ? (cellLeft + cellWidth) : -1f;
            List<Cell> allCells = CellsOf(table != null ? table.Range.Cells : null);

            Cell closestBelowCell = null;
            int closestBelowRow = int.MaxValue;
            int maxRowIndex = targetRow;
            var rowCells = new Dictionary<int, List<Cell>>();

            foreach (Cell c in allCells)
            {
                try
                {
                    int r = c.RowIndex;
                    if (r > maxRowIndex) maxRowIndex = r;
                    if (r <= targetRow) continue;
                    if (!rowCells.ContainsKey(r))
                        rowCells[r] = new List<Cell>();
                    rowCells[r].Add(c);
                }
                catch { }
            }

            // 普通单行格（nominalCellHeight > 0）：只需看下一行首个单元格的 TopY
            if (nominalCellHeight > 4f && rowCells.ContainsKey(targetRow + 1) && rowCells[targetRow + 1].Count > 0)
            {
                Cell nextRowFirst = rowCells[targetRow + 1][0];
                float nextTopY = MeasureCellTopYFast(nextRowFirst, topPad, bottomPad);
                if (nextTopY > cellTopY + 4f)
                    return Math.Max(nominalCellHeight, nextTopY - cellTopY);
                return nominalCellHeight;
            }

            // 纵向合并单元格（如跨 3 行的照片格）：按行寻找与 [cellLeft, cellRight] 水平重叠的最邻近下方单元格
            var sortedRows = new List<int>(rowCells.Keys);
            sortedRows.Sort();

            foreach (int r in sortedRows)
            {
                if (r >= closestBelowRow) break;

                foreach (Cell c in rowCells[r])
                {
                    try
                    {
                        int cStart = c.Range.Start;
                        Range cr = table.Range.Document.Range(cStart, cStart);
                        float y = GetPageY(cr);
                        if (y <= cellTopY + 3f) continue;

                        int col = c.ColumnIndex;
                        float cw = 0f;
                        try { cw = c.Width; } catch { }

                        float cx = GetPageX(cr);
                        float cLeft = Math.Max(0f, cx - leftPad);
                        float cRight = cLeft + cw;

                        bool overlapsHorizontally = false;
                        if (cellLeft >= 0f && cellRight > cellLeft && cw > 0f)
                        {
                            float overlapLeft = Math.Max(cellLeft, cLeft);
                            float overlapRight = Math.Min(cellRight, cRight);
                            if (overlapRight - overlapLeft > 4f)
                                overlapsHorizontally = true;
                        }
                        else if (col == targetCol)
                        {
                            overlapsHorizontally = true;
                        }

                        if (overlapsHorizontally)
                        {
                            closestBelowCell = c;
                            closestBelowRow = r;
                            break;
                        }
                    }
                    catch { }
                }
            }

            if (closestBelowCell != null)
            {
                float belowTopY = MeasureCellTopYFast(closestBelowCell, topPad, bottomPad);
                if (belowTopY > cellTopY + 4f)
                    return belowTopY - cellTopY;
            }

            if (nominalCellHeight > 4f)
                return nominalCellHeight;

            if (rowCells.ContainsKey(maxRowIndex))
            {
                foreach (Cell lastRowCell in rowCells[maxRowIndex])
                {
                    try
                    {
                        float lh = lastRowCell.Height;
                        if (lh > 4f && lh < 10000f)
                        {
                            float lastTop = MeasureCellTopYFast(lastRowCell, topPad, bottomPad);
                            if (lastTop >= cellTopY)
                                return (lastTop + lh) - cellTopY;
                        }
                    }
                    catch { }
                }
            }

            return 28f;
        }

        private static readonly Regex SignatureRegex = new Regex(
            @"(手签|签字|签名|签章|盖章|公章|负责人|审核人|审批人|申请人\s*[（\(]|年\s*月\s*日|\d{4}\s*年\s*\d{1,2}\s*月)",
            RegexOptions.Compiled);

        /// <summary>判断一段文本是否属于单元格底部的落款/签字/盖章/日期行。</summary>
        public static bool IsSignatureOrDateLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            string clean = CleanCellText(text).Trim();
            if (clean.Length == 0 || clean.Length > 45) return false;
            return SignatureRegex.IsMatch(clean);
        }
    }
}
