using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using NetOffice.OfficeApi.Enums;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Forms
{
    public sealed class PhotoFitResult
    {
        public bool Success;
        public bool FromClipboard;
        public string Message;
        public string TimingText;
        public Cell TargetCell;
        public Shape CreatedShape;
    }

    /// <summary>
    /// 证件照「一键严丝合缝入格」（Smart Photo Fit-to-Cell）。
    /// 支持三种自然路径：
    /// ① 已粘贴进单元格的 InlineShape / 已选中的浮动 Shape → 转悬浮、恢复被撑坏的行高、裁切入格；
    /// ② 光标在文档/单元格内 + 剪贴板有图片/文件 → 直接从剪贴板严丝合缝贴入；
    /// ③ 剪贴板无图 → 弹出文件选择框选图贴入。
    /// 全程包裹在单次 UndoRecord 内，清空单元格内“照片/免冠照”提示字（Ctrl+Z 一次还原）。
    /// </summary>
    public static class PhotoCellFitter
    {
        private const float BorderInsetPt = 0.75f;
        private const string ShapeTagPrefix = "TypeRight:Photo:";

        private static readonly Regex PhotoCellHintRegex = new Regex(
            @"(照\s*片|相\s*片|免\s*冠|寸\s*照|头像|Photo)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// 执行照片严丝合缝入格。
        /// 若 forceFilePath 不为空，则直接使用该本地图片文件贴入目标格。
        /// </summary>
        public static PhotoFitResult FitPhoto(Application app, string forceFilePath = null)
        {
            var result = new PhotoFitResult();
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
                result.Message = "请先打开需要贴照片的文档。";
                return result;
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();
            FormGeometry.EnsurePrintView(app);

            // 1. 识别是否已在文档中选中了图片（InlineShape 或 Shape）——先判 sel.Type 避免无图时访问 ShapeRange 抛 COM 异常耗时
            InlineShape selectedInline = null;
            Shape selectedShape = null;

            if (string.IsNullOrEmpty(forceFilePath))
            {
                try
                {
                    WdSelectionType selType = sel.Type;
                    if (selType == WdSelectionType.wdSelectionInlineShape)
                    {
                        var inlines = FormGeometry.InlineShapesOf(sel.InlineShapes);
                        if (inlines.Count > 0) selectedInline = inlines[0];
                    }
                    else if (selType == WdSelectionType.wdSelectionShape)
                    {
                        if (sel.ShapeRange != null && sel.ShapeRange.Count > 0)
                        {
                            foreach (object item in (System.Collections.IEnumerable)sel.ShapeRange.UnderlyingObject)
                            {
                                var sh = new Shape(null, item);
                                string alt = SafeAltText(sh);
                                if (alt.StartsWith("TypeRight:CellBox") || alt.StartsWith("TypeRight:UnderlineBox"))
                                    continue;
                                selectedShape = sh;
                                break;
                            }
                        }
                    }
                }
                catch { }
            }
            long tSel = sw.ElapsedMilliseconds;

            // 2. 定位目标表格与目标「照片单元格」
            Table table = null;
            Cell targetCell = null;

            try
            {
                if (sel.Tables != null && sel.Tables.Count > 0)
                    table = sel.Tables[1];
                else if (selectedInline != null && selectedInline.Range != null && selectedInline.Range.Tables.Count > 0)
                    table = selectedInline.Range.Tables[1];
                else if (selectedShape != null && selectedShape.Anchor != null && selectedShape.Anchor.Tables.Count > 0)
                    table = selectedShape.Anchor.Tables[1];
                else
                {
                    var allTables = WordCompat.TablesOf(doc.Tables);
                    if (allTables.Count > 0) table = allTables[0];
                }
            }
            catch { }

            if (table == null)
            {
                result.Message = "未找到表格，请先将光标点进需要贴照片的表格单元格内。";
                return result;
            }

            try
            {
                if (sel.Cells != null && sel.Cells.Count > 0)
                    targetCell = sel.Cells[1];
                else if (selectedInline != null && selectedInline.Range != null && selectedInline.Range.Cells.Count > 0)
                    targetCell = selectedInline.Range.Cells[1];
                else if (selectedShape != null && selectedShape.Anchor != null && selectedShape.Anchor.Cells.Count > 0)
                    targetCell = selectedShape.Anchor.Cells[1];
            }
            catch { }

            // 若用户已直接点在照片格内（含“照片/免冠”字样、或纵向合并格、或高度>=60pt、或格内已有图片），直接使用当前格，省去全表文本扫描！
            if (targetCell == null)
            {
                targetCell = FindExplicitPhotoCell(table) ?? FindLikelyPhotoCellByGeometry(table);
            }
            else
            {
                bool isAlreadyPhotoCell = false;
                try
                {
                    float ch = targetCell.Height;
                    if (ch >= 60f) isAlreadyPhotoCell = true; // 含纵向合并格(9999999)或高大照片格
                    else if (targetCell.ColumnIndex >= 5 && targetCell.RowIndex <= 2 && targetCell.Width >= 50f)
                        isAlreadyPhotoCell = true;
                }
                catch { }

                if (!isAlreadyPhotoCell)
                {
                    try { if (targetCell.Range.InlineShapes.Count > 0) isAlreadyPhotoCell = true; } catch { }
                }
                if (!isAlreadyPhotoCell)
                {
                    string curText = FormGeometry.CleanCellText(SafeCellText(targetCell));
                    if (PhotoCellHintRegex.IsMatch(curText)) isAlreadyPhotoCell = true;
                }

                if (!isAlreadyPhotoCell)
                {
                    Cell bestPhotoCell = FindExplicitPhotoCell(table) ?? FindLikelyPhotoCellByGeometry(table);
                    if (bestPhotoCell != null)
                        targetCell = bestPhotoCell;
                }
            }

            if (targetCell == null)
            {
                result.Message = "请先将光标点进需要放置照片的表格单元格。";
                return result;
            }

            // 若目标单元格内已经有 InlineShape（即使没被框选，只是光标在格子里），也自动纳入转换
            if (string.IsNullOrEmpty(forceFilePath) && selectedInline == null && selectedShape == null)
            {
                try
                {
                    var cellInlines = FormGeometry.InlineShapesOf(targetCell.Range.InlineShapes);
                    if (cellInlines.Count > 0) selectedInline = cellInlines[0];
                }
                catch { }
            }

            // 3. 若既未选中文档内图片，又未指定文件路径，则尝试读取剪贴板图片或弹出文件选择框
            //    注意：若直接使用本地图片文件（chosenFilePath），无需经过 GDI+ 裁切与 PNG 重编码落盘，直接交给 Word 原生加载，既保留 100% 完整画面又省去磁盘/编码耗时！
            Bitmap clipboardBitmap = null;
            bool fromClipboard = false;
            string chosenFilePath = forceFilePath;

            if (selectedInline == null && selectedShape == null)
            {
                if (string.IsNullOrEmpty(chosenFilePath) || !File.Exists(chosenFilePath))
                {
                    string clipFile;
                    clipboardBitmap = TryGetBitmapOrFileFromClipboard(out clipFile);
                    if (!string.IsNullOrEmpty(clipFile))
                    {
                        chosenFilePath = clipFile;
                        fromClipboard = true;
                    }
                    else if (clipboardBitmap != null)
                    {
                        fromClipboard = true;
                    }
                    else
                    {
                        // 剪贴板无图 → 弹出文件选择框
                        chosenFilePath = PromptSelectImageFile();
                        if (string.IsNullOrEmpty(chosenFilePath))
                        {
                            result.Message = ""; // 用户主动取消文件对话框，静默返回
                            return result;
                        }
                        if (!File.Exists(chosenFilePath))
                        {
                            result.Message = "未找到所选图片文件。";
                            return result;
                        }
                    }
                }
            }

            long tCell = sw.ElapsedMilliseconds;

            // 4. 在单个 UndoRecord + ScreenUpdating=false 内执行：
            //    ① 清空提示字并归零单元格顶端对齐 → 测定格子宽高 cellW/cellH 与宽高比 cellRatio；
            //    ② 读取图片宽高 imgW/imgH 与宽高比 imgRatio → 先等比例缩放对齐，若宽高比不一致再微调尺寸做到四边严丝合缝！
            UndoRecord undo = WordCompat.BeginUndo(app, "TypeRight 照片严丝合缝入格");
            int savedScrollV = -1;
            try { savedScrollV = app.ActiveWindow.VerticalPercentScrolled; } catch { }
            try { app.ScreenUpdating = false; } catch { }
            string tempPng = null;
            try
            {
                int rowIdx = 1, colIdx = 1;
                try { rowIdx = targetCell.RowIndex; colIdx = targetCell.ColumnIndex; } catch { }
                string cellTag = ShapeTagPrefix + rowIdx + "_" + colIdx;

                // 4.1 清理同一单元格上此前由 TypeRight 贴入的旧照片（防叠罗汉 E5）
                RemoveExistingTaggedPhoto(doc, cellTag, selectedShape);

                // 4.2 若是单元格内的 InlineShape，先转为浮于文字上方的 Shape，让被撑大的表格行立即回弹复原
                if (selectedInline != null)
                {
                    try
                    {
                        selectedShape = selectedInline.ConvertToShape();
                        selectedShape.WrapFormat.Type = WdWrapType.wdWrapFront;
                    }
                    catch { }
                }
                else if (selectedShape != null)
                {
                    try { selectedShape.WrapFormat.Type = WdWrapType.wdWrapFront; } catch { }
                }

                // 4.3 清空单元格内的提示文字（如“近期一寸免冠照片”）并设为顶端对齐
                ClearCellPromptTextAndNormalize(doc, targetCell);
                long tClear = sw.ElapsedMilliseconds;

                // 4.4 第 ① 步：测定当前格子的宽、高及宽高比
                float cellW, cellH;
                FormGeometry.MeasurePhotoCellSizeAfterNormalize(table, targetCell, out cellW, out cellH);
                long tMeasure = sw.ElapsedMilliseconds;
                if (cellW <= 4f || cellH <= 4f)
                {
                    result.Message = "无法测定当前单元格尺寸，请确认文档处于页面视图。";
                    return result;
                }

                float targetW = Math.Max(10f, cellW - BorderInsetPt * 2f);
                float targetH = Math.Max(10f, cellH - BorderInsetPt * 2f);

                string imageFileToInsert = chosenFilePath;
                if (string.IsNullOrEmpty(imageFileToInsert) && clipboardBitmap != null)
                {
                    try
                    {
                        using (clipboardBitmap)
                        {
                            tempPng = Path.Combine(Path.GetTempPath(), "TypeRight_Photo_" + Guid.NewGuid().ToString("N") + ".png");
                            clipboardBitmap.Save(tempPng, ImageFormat.Png);
                            imageFileToInsert = tempPng;
                        }
                        clipboardBitmap = null;
                    }
                    catch (Exception ex)
                    {
                        result.Message = "处理剪贴板图像失败：" + ex.Message;
                        return result;
                    }
                }
                long tCrop = sw.ElapsedMilliseconds;

                Range anchorRng = doc.Range(targetCell.Range.Start, targetCell.Range.Start);

                if (!string.IsNullOrEmpty(imageFileToInsert) && File.Exists(imageFileToInsert))
                {
                    object linkToFile = false;
                    object saveWithDoc = true;
                    object leftObj = BorderInsetPt;
                    object topObj = BorderInsetPt;
                    object missing = System.Type.Missing;
                    object anchorObj = anchorRng.UnderlyingObject;

                    // 先按原图原始尺寸与比例插入，再执行「① 对比格子与图片宽高比 → ② 先等比缩放，比例不一致时微调尺寸严丝合缝」
                    Shape picShape = doc.Shapes.AddPicture(
                        imageFileToInsert, linkToFile, saveWithDoc, leftObj, topObj, missing, missing, anchorObj);
                    long tAddPic = sw.ElapsedMilliseconds;

                    ScaleAndAlignPhotoShape(picShape, targetW, targetH, BorderInsetPt, BorderInsetPt, cellTag);
                    long tCfg = sw.ElapsedMilliseconds;
                    result.TimingText = string.Format("sel={0} cell={1} clear={2} measure={3} prep={4} addPic={5} cfg={6}",
                        tSel, tCell - tSel, tClear - tCell, tMeasure - tClear, tCrop - tMeasure, tAddPic - tCrop, tCfg - tAddPic);
                    result.CreatedShape = picShape;
                }
                else if (selectedShape != null)
                {
                    FormGeometry.RectPt cellRect = new FormGeometry.RectPt { Left = 0f, Top = 0f, Width = cellW, Height = cellH, Page = 1 };
                    float targetLeft = BorderInsetPt;
                    float targetTop = BorderInsetPt;
                    if (!FormGeometry.IsShapeLayoutInCell(selectedShape))
                    {
                        cellRect = FormGeometry.MeasureCellRect(table, targetCell, normalizeInPlace: false);
                        targetLeft = cellRect.Left + BorderInsetPt;
                        targetTop = cellRect.Top + BorderInsetPt;
                    }
                    FitExistingShapeToRect(selectedShape, table, targetCell, cellRect, targetLeft, targetTop, targetW, targetH, cellTag);
                    result.CreatedShape = selectedShape;
                }

                if (savedScrollV >= 0)
                {
                    try
                    {
                        if (app.ActiveWindow.VerticalPercentScrolled != savedScrollV)
                            app.ActiveWindow.VerticalPercentScrolled = savedScrollV;
                    }
                    catch { }
                }

                result.Success = true;
                result.FromClipboard = fromClipboard;
                result.TargetCell = targetCell;
                result.Message = fromClipboard
                    ? "✓ 已将剪贴板照片严丝合缝贴入单元格（Ctrl+Z 可撤销）"
                    : "✓ 照片已严丝合缝贴入单元格（Ctrl+Z 可撤销）";
                return result;
            }
            catch (Exception ex)
            {
                result.Message = "照片入格失败：" + ex.Message;
                return result;
            }
            finally
            {
                try { app.ScreenUpdating = true; } catch { }
                if (!string.IsNullOrEmpty(tempPng))
                {
                    try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
                }
                if (clipboardBitmap != null)
                {
                    try { clipboardBitmap.Dispose(); } catch { }
                }
                WordCompat.EndUndo(undo);
            }
        }

        /// <summary>弹出文件选择对话框供用户挑选证件照。</summary>
        public static string PromptSelectImageFile()
        {
            try
            {
                using (var dlg = new System.Windows.Forms.OpenFileDialog())
                {
                    dlg.Title = "选择要贴入单元格的证件照 / 图片";
                    dlg.Filter = "图片文件 (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|所有文件 (*.*)|*.*";
                    dlg.Multiselect = false;
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        return dlg.FileName;
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 辅助方法：按目标宽高比裁切图像（保留供单元测试使用）。
        /// </summary>
        public static Bitmap CropBitmapToAspect(Bitmap src, float targetAspect)
        {
            if (src == null) return null;
            int w = src.Width;
            int h = src.Height;
            if (w <= 4 || h <= 4 || targetAspect <= 0.05f)
                return new Bitmap(src);

            float srcAspect = (float)w / h;
            System.Drawing.Rectangle cropRect;

            if (Math.Abs(srcAspect - targetAspect) < 0.01f)
            {
                cropRect = new System.Drawing.Rectangle(0, 0, w, h);
            }
            else if (srcAspect > targetAspect)
            {
                int newW = Math.Max(1, (int)Math.Round(h * targetAspect));
                int cropX = Math.Max(0, (w - newW) / 2);
                cropRect = new System.Drawing.Rectangle(cropX, 0, Math.Min(newW, w - cropX), h);
            }
            else
            {
                int newH = Math.Max(1, (int)Math.Round(w / targetAspect));
                int overflowH = Math.Max(0, h - newH);
                int cropY = Math.Max(0, (int)Math.Round(overflowH * 0.15));
                cropRect = new System.Drawing.Rectangle(0, cropY, w, Math.Min(newH, h - cropY));
            }

            int dstW = cropRect.Width;
            int dstH = cropRect.Height;
            const int MaxPhotoDim = 640;
            if (dstW > MaxPhotoDim || dstH > MaxPhotoDim)
            {
                float scale = (float)MaxPhotoDim / Math.Max(dstW, dstH);
                dstW = Math.Max(1, (int)Math.Round(dstW * scale));
                dstH = Math.Max(1, (int)Math.Round(dstH * scale));
            }

            var dst = new Bitmap(dstW, dstH, PixelFormat.Format32bppArgb);
            dst.SetResolution(96f, 96f);

            using (Graphics g = Graphics.FromImage(dst))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.DrawImage(src, new System.Drawing.Rectangle(0, 0, dst.Width, dst.Height), cropRect, GraphicsUnit.Pixel);
            }
            return dst;
        }

        /// <summary>
        /// 按两步法将图片严丝合缝放入目标单元格：
        /// ① 读取当前图片的宽 imgW、高 imgH、宽高比 imgRatio，与目标格宽 targetW、高 targetH、宽高比 cellRatio 对比；
        /// ② 先开启等比例缩放（LockAspectRatio = msoTrue）将图片等比缩放到与格子宽度对齐；
        ///   - 若原本宽高比一致（|imgRatio - cellRatio| &lt;= 0.01），直接移入格子左上角即完成；
        ///   - 若宽高比不一致，解锁纵横比（LockAspectRatio = msoFalse）微调图片尺寸至 (targetW, targetH)，实现四边严丝合缝对齐且不裁切任何画面内容。
        /// </summary>
        private static void ScaleAndAlignPhotoShape(Shape shape, float targetW, float targetH, float targetLeft, float targetTop, string tag)
        {
            if (shape == null) return;

            // 显式指定相对于 Page + LayoutInCell，使 (0,0) 原点精确落在单元格四条外边框上（彻底消除 LeftPadding=5.4pt 右偏与段落上偏！）
            try { shape.WrapFormat.Type = WdWrapType.wdWrapFront; } catch { }
            try { shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage; } catch { }
            try { shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage; } catch { }

            float imgW = 0f, imgH = 0f;
            try { imgW = shape.Width; imgH = shape.Height; } catch { }

            float cellRatio = targetH > 1f ? (targetW / targetH) : 0.7f;
            float imgRatio = (imgW > 1f && imgH > 1f) ? (imgW / imgH) : cellRatio;

            // 第一步：初始调控先等比例缩放至宽度对齐
            try
            {
                shape.LockAspectRatio = MsoTriState.msoTrue;
                shape.Width = targetW;
            }
            catch { }

            // 第二步：若图片与格子的宽高比不一致，解锁比例微调图片尺寸以做到四边严丝合缝对齐
            if (Math.Abs(imgRatio - cellRatio) > 0.008f)
            {
                try
                {
                    shape.LockAspectRatio = MsoTriState.msoFalse;
                    shape.Width = targetW;
                    shape.Height = targetH;
                }
                catch { }
            }

            try { shape.Left = targetLeft; } catch { }
            try { shape.Top = targetTop; } catch { }
            try { shape.AlternativeText = tag; } catch { }
        }

        /// <summary>对文档中已有的 Shape 应用「等比缩放 + 尺寸微调严丝合缝定位」。</summary>
        private static void FitExistingShapeToRect(
            Shape shape,
            Table table,
            Cell targetCell,
            FormGeometry.RectPt targetCellRect,
            float targetLeft,
            float targetTop,
            float targetW,
            float targetH,
            string tag)
        {
            if (shape == null) return;
            try
            {
                shape.PictureFormat.CropLeft = 0f;
                shape.PictureFormat.CropRight = 0f;
                shape.PictureFormat.CropTop = 0f;
                shape.PictureFormat.CropBottom = 0f;
            }
            catch { }

            ScaleAndAlignPhotoShape(shape, targetW, targetH, targetLeft, targetTop, tag);
        }

        /// <summary>清空单元格内的文字（保留单元格结束标记 \r\a，不破坏表格结构）并归零对齐。</summary>
        private static void ClearCellPromptTextAndNormalize(Document doc, Cell cell)
        {
            try
            {
                int start = cell.Range.Start;
                int end = cell.Range.End - 1;
                if (end > start)
                {
                    Range textRng = doc.Range(start, end);
                    textRng.Text = "";
                }
            }
            catch { }
            try
            {
                if (cell.VerticalAlignment != WdCellVerticalAlignment.wdCellAlignVerticalTop)
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
            }
            catch { }
            try
            {
                Paragraph p1 = cell.Range.Paragraphs[1];
                if (p1.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                    p1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            }
            catch { }
        }

        private static void RemoveExistingTaggedPhoto(Document doc, string cellTag, Shape keepShape)
        {
            try
            {
                foreach (Shape sh in FormGeometry.ShapesOf(doc.Shapes))
                {
                    if (keepShape != null && ReferenceEquals(sh.UnderlyingObject, keepShape.UnderlyingObject))
                        continue;
                    if (SafeAltText(sh) == cellTag)
                    {
                        try { sh.Delete(); } catch { }
                    }
                }
            }
            catch { }
        }

        private static Bitmap TryGetBitmapOrFileFromClipboard(out string filePath)
        {
            filePath = null;
            try
            {
                if (System.Windows.Forms.Clipboard.ContainsFileDropList())
                {
                    var files = System.Windows.Forms.Clipboard.GetFileDropList();
                    foreach (string f in files)
                    {
                        if (string.IsNullOrEmpty(f) || !File.Exists(f)) continue;
                        string ext = Path.GetExtension(f).ToLowerInvariant();
                        if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp")
                        {
                            filePath = f;
                            return null;
                        }
                    }
                }
                if (System.Windows.Forms.Clipboard.ContainsImage())
                {
                    Image img = System.Windows.Forms.Clipboard.GetImage();
                    if (img != null) return new Bitmap(img);
                }
            }
            catch { }
            return null;
        }

        /// <summary>在表格中查找写有“照片/免冠/寸照”等明确提示字的单元格。</summary>
        private static Cell FindExplicitPhotoCell(Table table)
        {
            foreach (Cell c in FormGeometry.CellsOf(table.Range.Cells))
            {
                string txt = FormGeometry.CleanCellText(SafeCellText(c));
                if (txt.Length > 0 && txt.Length <= 24 && PhotoCellHintRegex.IsMatch(txt))
                    return c;
            }
            return null;
        }

        /// <summary>按几何特征查找表格右上角的竖长合并照片格（起始行<=2，位于最右列）。</summary>
        private static Cell FindLikelyPhotoCellByGeometry(Table table)
        {
            Cell best = null;
            int maxCol = 0;
            var cells = FormGeometry.CellsOf(table.Range.Cells);
            foreach (Cell c in cells)
            {
                try
                {
                    int r = c.RowIndex;
                    int col = c.ColumnIndex;
                    if (r == 1 && col > maxCol) maxCol = col;
                }
                catch { }
            }
            foreach (Cell c in cells)
            {
                try
                {
                    int r = c.RowIndex;
                    int col = c.ColumnIndex;
                    if (r <= 2 && col >= Math.Max(2, maxCol))
                    {
                        string txt = FormGeometry.CleanCellText(SafeCellText(c));
                        if (txt.Length == 0 || PhotoCellHintRegex.IsMatch(txt))
                        {
                            best = c;
                            break;
                        }
                    }
                }
                catch { }
            }
            return best;
        }

        private static bool IsSameCell(Cell a, Cell b)
        {
            if (a == null || b == null) return false;
            try { return a.RowIndex == b.RowIndex && a.ColumnIndex == b.ColumnIndex; }
            catch { return false; }
        }

        private static string SafeCellText(Cell c)
        {
            try { return c.Range.Text ?? ""; }
            catch { return ""; }
        }

        private static string SafeAltText(Shape sh)
        {
            try { return sh.AlternativeText ?? ""; }
            catch { return ""; }
        }
    }
}
