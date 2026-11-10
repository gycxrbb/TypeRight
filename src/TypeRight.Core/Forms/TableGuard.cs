using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Forms
{
    public sealed class TableGuardResult
    {
        public bool Success;
        public int LockedRows;
        public int ShrunkCells;
        public bool CompressedTrailingBlank;
        public bool PulledBackToSinglePage;
        public string Message;
    }

    /// <summary>
    /// 能力一：表格锁死与死守一页（Table Lock &amp; Auto-Fit Shrink）。
    /// 1. 禁止表格自动重调列宽（AllowAutoFit = false）；
    /// 2. 识别被多字折行撑高的单行字段格并执行「三阶无感缩字」；
    /// 3. 实测每行真实视觉高度后锁死为固定行高（wdRowHeightExactly），防止后续打字撑破；
    /// 4. 压缩表格末尾删不掉的幽灵空段落（1pt 行距），消除空白第 2 页。
    /// </summary>
    public static class TableGuard
    {
        private const float MinFontSizePt = 7.5f;

        public static TableGuardResult LockAndGuardTable(Application app)
        {
            var result = new TableGuardResult();
            if (app == null) { result.Message = "未连接到 Word 应用程序。"; return result; }

            Document doc = null;
            Selection sel = null;
            try
            {
                doc = app.ActiveDocument;
                sel = app.Selection;
            }
            catch { }
            if (doc == null || sel == null)
            {
                result.Message = "请先打开需要锁定的表格文档。";
                return result;
            }

            FormGeometry.EnsurePrintView(app);

            Table table = null;
            try
            {
                if (sel.Tables != null && sel.Tables.Count > 0)
                    table = sel.Tables[1];
                else
                {
                    var tables = WordCompat.TablesOf(doc.Tables);
                    if (tables.Count > 0) table = tables[0];
                }
            }
            catch { }

            if (table == null)
            {
                result.Message = "文档中未找到表格。";
                return result;
            }

            UndoRecord undo = WordCompat.BeginUndo(app, "TypeRight 锁表防崩");
            try
            {
                int pageCountBefore = SafePageCount(doc);

                // 1. 锁死列宽自动重调
                try { table.AllowAutoFit = false; } catch { }

                List<Cell> allCells = FormGeometry.CellsOf(table.Range.Cells);
                List<Row> allRows = FormGeometry.RowsOf(table.Rows);

                // 2. 第一遍扫描单元格：取消网格对齐造成的虚胖行高，对短字段被意外折成多行的单元格先行缩字收拢
                int shrunkCount = 0;
                foreach (Cell cell in allCells)
                {
                    try
                    {
                        // 跳过包含嵌套子表的复合格（E7）
                        if (cell.Tables != null && cell.Tables.Count > 0) continue;

                        string txt = FormGeometry.CleanCellText(cell.Range.Text);
                        if (string.IsNullOrEmpty(txt))
                        {
                            NormalizeCellGrid(cell);
                            continue;
                        }

                        NormalizeCellGrid(cell);

                        // 判断是否为“短字段填写格”（单段落、字数 4~36 字，但因格子窄折成了 2~3 行）
                        List<Paragraph> paras = WordCompat.ParagraphsOf(cell.Range.Paragraphs);
                        if (paras.Count == 1 && txt.Length >= 4 && txt.Length <= 36 && !txt.Contains("\r") && !txt.Contains("\n"))
                        {
                            float cellW = 100f;
                            try { cellW = cell.Width; } catch { }
                            // 只针对中小型填写格（宽度 < 220pt）且非左侧 2~4 字短标签格
                            if (cellW < 220f && IsWrappedMultiLine(cell))
                            {
                                if (ShrinkSingleLineCellToFit(doc, cell))
                                    shrunkCount++;
                            }
                        }
                    }
                    catch { }
                }

                // 3. 实测每一行的当前真实视觉高度，并锁死为 wdRowHeightExactly（防 E3 双行表头被截断）
                var rowTopMap = MeasureRowTopCoordinates(allCells, table);
                int lockedRows = 0;

                for (int i = 0; i < allRows.Count; i++)
                {
                    Row row = allRows[i];
                    int rIdx = i + 1;
                    try
                    {
                        row.AllowBreakAcrossPages = 0;
                        if (row.HeightRule == WdRowHeightRule.wdRowHeightExactly && row.Height > 8f && row.Height < 1000f)
                        {
                            lockedRows++;
                            continue;
                        }

                        float measuredH = EstimateRowHeight(row, rIdx, rowTopMap, table);
                        if (measuredH >= 10f && measuredH < 750f)
                        {
                            row.Height = measuredH;
                            row.HeightRule = WdRowHeightRule.wdRowHeightExactly;
                            lockedRows++;
                        }
                    }
                    catch { }
                }

                // 4. 第二遍扫描：锁死固定行高后，检查是否有单元格内容超出底边被截断，执行三阶缩字
                rowTopMap = MeasureRowTopCoordinates(allCells, table);
                foreach (Cell cell in allCells)
                {
                    try
                    {
                        if (cell.Tables != null && cell.Tables.Count > 0) continue;
                        string txt = FormGeometry.CleanCellText(cell.Range.Text);
                        if (txt.Length < 2) continue;

                        if (IsCellContentClippedOrOverflowing(doc, table, cell, rowTopMap))
                        {
                            if (ShrinkOverflowingCell(doc, table, cell, rowTopMap))
                                shrunkCount++;
                        }
                    }
                    catch { }
                }

                // 5. 消除表格末尾删不掉的幽灵空白段落（解决表格占满 1 页时空白第 2 页问题）
                bool compressedTail = CompressTrailingBlankParagraphs(doc, table);

                // 6. 若原文档是 2 页且第 2 页仅有表格极少部分溢出（<= 55pt），微缩各行固定高度 4%~8% 拉回第 1 页
                bool pulledBack = false;
                int pageCountAfter = SafePageCount(doc);
                if (pageCountAfter == 2 && TryPullSlightlySpilledTableToOnePage(doc, table, allRows))
                {
                    pulledBack = true;
                    CompressTrailingBlankParagraphs(doc, table);
                }

                result.Success = true;
                result.LockedRows = lockedRows;
                result.ShrunkCells = shrunkCount;
                result.CompressedTrailingBlank = compressedTail;
                result.PulledBackToSinglePage = pulledBack || (pageCountBefore > 1 && SafePageCount(doc) == 1);

                var parts = new List<string>();
                parts.Add("✓ 已锁定 " + lockedRows + " 行固定高度与列宽");
                if (shrunkCount > 0) parts.Add("智能缩字适应 " + shrunkCount + " 格");
                if (result.PulledBackToSinglePage) parts.Add("已收拢死守 1 页");
                else if (compressedTail) parts.Add("已消除表尾冗余空行");
                parts.Add("Ctrl+Z 可撤销");

                result.Message = string.Join(" · ", parts);
                return result;
            }
            catch (Exception ex)
            {
                result.Message = "锁定表格失败：" + ex.Message;
                return result;
            }
            finally
            {
                WordCompat.EndUndo(undo);
            }
        }

        private static void NormalizeCellGrid(Cell cell)
        {
            try
            {
                ParagraphFormat pf = cell.Range.ParagraphFormat;
                pf.DisableLineHeightGrid = -1;
            }
            catch { }
        }

        /// <summary>判断单段单元格的首字符与尾字符是否落在了不同视觉行上（即发生了自动折行）。</summary>
        private static bool IsWrappedMultiLine(Cell cell)
        {
            try
            {
                Document doc = cell.Range.Document;
                int start = cell.Range.Start;
                int end = cell.Range.End - 1;
                if (end - start <= 1) return false;

                Range firstChar = doc.Range(start, start + 1);
                Range lastChar = doc.Range(end - 1, end);

                float y1 = FormGeometry.GetPageY(firstChar);
                float y2 = FormGeometry.GetPageY(lastChar);
                if (y1 < 0f || y2 < 0f) return false;

                return (y2 - y1) > 4.5f;
            }
            catch { return false; }
        }

        /// <summary>
        /// 对小格子（如“专业”“奖项名称”）执行三阶无感压缩，尝试把折成 2 行的单句收拢回 1 行或紧凑状态。
        /// </summary>
        private static bool ShrinkSingleLineCellToFit(Document doc, Cell cell)
        {
            try
            {
                int start = cell.Range.Start;
                int end = cell.Range.End - 1;
                if (end <= start) return false;

                Range rng = doc.Range(start, end);
                float origSize = 10.5f;
                try
                {
                    if (rng.Font.Size > 5f && rng.Font.Size < 72f)
                        origSize = rng.Font.Size;
                }
                catch { }

                // 阶段 1：清零段前段后距 + 紧缩字符间距（-0.4pt ~ -0.7pt）+ 缩减左右单元格边距
                try
                {
                    rng.ParagraphFormat.SpaceBefore = 0f;
                    rng.ParagraphFormat.SpaceAfter = 0f;
                    rng.ParagraphFormat.DisableLineHeightGrid = -1;
                    rng.Font.Spacing = -0.5f;
                }
                catch { }

                if (!IsWrappedMultiLine(cell)) return true;

                // 阶段 2：按 0.5pt 阶梯微降字号（最多降到 7.5pt）
                for (float sz = origSize - 0.5f; sz >= MinFontSizePt; sz -= 0.5f)
                {
                    try { rng.Font.Size = sz; } catch { break; }
                    if (!IsWrappedMultiLine(cell)) return true;
                }

                // 若降到 7.5pt 仍需 2 行（说明字数确实较多，需 2 行排布），将行距设为紧凑 0.92 倍并保持可读字号 8.5pt
                try
                {
                    rng.Font.Size = Math.Max(8.5f, origSize - 1.5f);
                    rng.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                    rng.ParagraphFormat.LineSpacing = rng.Font.Size * 1.08f;
                }
                catch { }
                return true;
            }
            catch { return false; }
        }

        /// <summary>判断固定行高下的单元格是否发生了内容溢出/底部截断。</summary>
        private static bool IsCellContentClippedOrOverflowing(Document doc, Table table, Cell cell, SortedDictionary<int, RowCoord> rowTopMap)
        {
            try
            {
                int start = cell.Range.Start;
                int end = cell.Range.End - 1;
                if (end - start <= 1) return false;

                // 若底部是签字落款格且有很多空行，先看最后一个非空字符
                Range lastChar = doc.Range(end - 1, end);
                float lastY = FormGeometry.GetPageY(lastChar);

                // 在 wdRowHeightExactly 下，若末字符已被截断不可见，Word 的 Information(wdVerticalPositionRelativeToPage) 常返回 -1
                if (lastY < 0f) return true;

                FormGeometry.RectPt rect = FormGeometry.MeasureCellRect(table, cell, normalizeInPlace: false);
                if (!rect.IsValid) return false;

                float fontSize = 10.5f;
                try { if (lastChar.Font.Size > 5f && lastChar.Font.Size < 72f) fontSize = lastChar.Font.Size; } catch { }

                // 末字符底部（lastY + 字号）若超过或紧贴单元格底边，则判定为溢出风险
                return (lastY + fontSize * 0.95f) > rect.Bottom;
            }
            catch { return false; }
        }

        /// <summary>对多行溢出单元格执行三阶缩字（去段间距/紧缩行距 → 0.5pt 阶梯降字号）。</summary>
        private static bool ShrinkOverflowingCell(Document doc, Table table, Cell cell, SortedDictionary<int, RowCoord> rowTopMap)
        {
            try
            {
                int start = cell.Range.Start;
                int end = cell.Range.End - 1;
                if (end <= start) return false;

                Range rng = doc.Range(start, end);
                float origSize = 10.5f;
                try { if (rng.Font.Size > 5f && rng.Font.Size < 72f) origSize = rng.Font.Size; } catch { }

                // 阶段 1：消肿（去网格对齐、清段间距、紧缩字符间距与行距）
                try
                {
                    rng.ParagraphFormat.DisableLineHeightGrid = -1;
                    rng.ParagraphFormat.SpaceBefore = 0f;
                    rng.ParagraphFormat.SpaceAfter = 0f;
                    rng.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                    rng.ParagraphFormat.LineSpacing = origSize * 1.12f;
                    rng.Font.Spacing = -0.4f;
                }
                catch { }

                if (!IsCellContentClippedOrOverflowing(doc, table, cell, rowTopMap))
                    return true;

                // 阶段 2：0.5pt 阶梯降字号（下限 7.5pt）
                for (float sz = origSize - 0.5f; sz >= MinFontSizePt; sz -= 0.5f)
                {
                    try
                    {
                        rng.Font.Size = sz;
                        rng.ParagraphFormat.LineSpacing = sz * 1.10f;
                    }
                    catch { break; }

                    if (!IsCellContentClippedOrOverflowing(doc, table, cell, rowTopMap))
                        return true;
                }

                return true;
            }
            catch { return false; }
        }

        private struct RowCoord
        {
            public float TopY;
            public int Page;
        }

        private static SortedDictionary<int, RowCoord> MeasureRowTopCoordinates(List<Cell> allCells, Table table)
        {
            var map = new SortedDictionary<int, RowCoord>();
            float topPad = 0f;
            try { topPad = table.TopPadding; } catch { }

            foreach (Cell c in allCells)
            {
                try
                {
                    int r = c.RowIndex;
                    Range cr = c.Range.Duplicate;
                    cr.Collapse(WdCollapseDirection.wdCollapseStart);
                    float y = FormGeometry.GetPageY(cr);
                    int pg = FormGeometry.GetPageNumber(cr);
                    if (y < 0f) continue;

                    float rowTop = Math.Max(0f, y - topPad);
                    if (!map.ContainsKey(r) || rowTop < map[r].TopY)
                    {
                        map[r] = new RowCoord { TopY = rowTop, Page = pg };
                    }
                }
                catch { }
            }
            return map;
        }

        private static float EstimateRowHeight(Row row, int rIdx, SortedDictionary<int, RowCoord> rowTopMap, Table table)
        {
            // 优先使用相邻两行的页面实测 TopY 之差（防 E3 双行表头坍缩 & 防 E12 跨页负值）
            if (rowTopMap.ContainsKey(rIdx) && rowTopMap.ContainsKey(rIdx + 1))
            {
                RowCoord cur = rowTopMap[rIdx];
                RowCoord next = rowTopMap[rIdx + 1];
                if (cur.Page == next.Page && next.TopY > cur.TopY + 6f)
                    return next.TopY - cur.TopY;
            }

            // 最后一行：尝试读取表格后紧跟段落的 TopY
            if (rowTopMap.ContainsKey(rIdx) && table != null)
            {
                try
                {
                    Document doc = table.Range.Document;
                    int afterPos = table.Range.End;
                    if (afterPos < doc.Content.End)
                    {
                        Range afterRng = doc.Range(afterPos, afterPos);
                        int afterPage = FormGeometry.GetPageNumber(afterRng);
                        float afterY = FormGeometry.GetPageY(afterRng);
                        RowCoord cur = rowTopMap[rIdx];
                        if (afterPage == cur.Page && afterY > cur.TopY + 8f)
                            return afterY - cur.TopY;
                    }
                }
                catch { }
            }

            // 兜底读取 Row.Height（若非未定义 9999999）
            try
            {
                float h = row.Height;
                if (h >= 10f && h < 800f) return h;
            }
            catch { }

            return 24f;
        }

        /// <summary>压缩表格紧邻下方的空白段落为 1pt 固定行距，消除占满一页时的空白第 2 页。</summary>
        private static bool CompressTrailingBlankParagraphs(Document doc, Table table)
        {
            bool changed = false;
            try
            {
                int pos = table.Range.End;
                int docEnd = doc.Content.End;
                if (pos >= docEnd) return false;

                Range tailRng = doc.Range(pos, docEnd);
                List<Paragraph> tailParas = WordCompat.ParagraphsOf(tailRng.Paragraphs);
                foreach (Paragraph p in tailParas)
                {
                    // 若遇到另一个表格，立即停止
                    try { if (p.Range.Tables != null && p.Range.Tables.Count > 0) break; } catch { }

                    string txt = FormGeometry.CleanCellText(p.Range.Text);
                    if (string.IsNullOrWhiteSpace(txt))
                    {
                        p.Range.Font.Size = 1f;
                        p.Format.SpaceBefore = 0f;
                        p.Format.SpaceAfter = 0f;
                        p.Format.DisableLineHeightGrid = -1;
                        p.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                        p.Format.LineSpacing = 1f;
                        changed = true;
                    }
                    else
                    {
                        // 若下方是简短的“注：此表一式两份……”说明文字，收紧其段前段后距帮助留在第 1 页
                        if (txt.Length <= 80)
                        {
                            p.Format.SpaceBefore = 0f;
                            p.Format.SpaceAfter = 0f;
                            p.Format.DisableLineHeightGrid = -1;
                        }
                        break;
                    }
                }
            }
            catch { }
            return changed;
        }

        /// <summary>当表格因为细微膨胀刚巧溢出到第 2 页一点点时，按比例微缩大行高将其拉回第 1 页。</summary>
        private static bool TryPullSlightlySpilledTableToOnePage(Document doc, Table table, List<Row> allRows)
        {
            try
            {
                Range startRng = doc.Range(table.Range.Start, table.Range.Start);
                Range endRng = doc.Range(table.Range.End - 1, table.Range.End - 1);
                if (FormGeometry.GetPageNumber(startRng) != 1 || FormGeometry.GetPageNumber(endRng) != 2)
                    return false;

                // 尝试两轮温和微缩（每轮对高度 > 24pt 的行缩减 4%），看能否回到 1 页
                for (int pass = 0; pass < 2; pass++)
                {
                    foreach (Row r in allRows)
                    {
                        try
                        {
                            if (r.HeightRule == WdRowHeightRule.wdRowHeightExactly && r.Height > 24f && r.Height < 700f)
                            {
                                r.Height = Math.Max(18f, r.Height * 0.96f);
                            }
                        }
                        catch { }
                    }
                    if (SafePageCount(doc) == 1) return true;
                }
            }
            catch { }
            return false;
        }

        private static int SafePageCount(Document doc)
        {
            try
            {
                Range endRng = doc.Range(Math.Max(0, doc.Content.End - 1), doc.Content.End);
                return FormGeometry.GetPageNumber(endRng);
            }
            catch { return 1; }
        }
    }
}
