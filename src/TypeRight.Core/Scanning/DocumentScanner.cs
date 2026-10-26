using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Classification;
using TypeRight.Formatting;
using TypeRight.Formulas;

namespace TypeRight.Scanning
{
    /// <summary>扫描结果：段落特征 + 文档级统计。</summary>
    public sealed class ScanResult
    {
        public List<ParagraphFeatures> Features = new List<ParagraphFeatures>();
        public int TableCount;
        public TimeSpan Duration;
    }

    /// <summary>
    /// 段落/表格遍历与特征提取（PRD §4.2 第一段）。doc.Paragraphs 包含表格内段落。
    /// </summary>
    public sealed class DocumentScanner
    {
        // 编号模式（PRD §4.2 特征表）
        private static readonly Regex PatternChapter = new Regex(
            @"^(第[一二三四五六七八九十百\d]+[章节部分篇]|模块[一二三四五六七八九十百\d]+|Chapter\s+\d+)", RegexOptions.Compiled);
        private static readonly Regex PatternNumeric3 = new Regex(@"^\d+\.\d+\.\d+[\s．.、]", RegexOptions.Compiled);
        private static readonly Regex PatternNumeric2 = new Regex(@"^\d+\.\d+[\s．.、]", RegexOptions.Compiled);
        private static readonly Regex PatternChinese = new Regex(@"^[一二三四五六七八九十]+、", RegexOptions.Compiled);
        private static readonly Regex PatternCircled = new Regex(@"^[①-⑳]", RegexOptions.Compiled);
        private static readonly Regex PatternBracket = new Regex(@"^【[^】]{0,30}】", RegexOptions.Compiled);
        private static readonly Regex PatternCaption = new Regex(
            @"^(图|表|Figure|Table|【图|【表)\s*\S", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private readonly Application _app;

        public DocumentScanner(Application app)
        {
            if (app == null) throw new ArgumentNullException("app");
            _app = app;
        }

        /// <summary>扫描整个文档；restrictTo 选区时仅扫描选区内段落（用于选区模式）。</summary>
        public ScanResult Scan(Document doc, Range restrictTo = null)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var result = new ScanResult();
            try { result.TableCount = doc.Tables.Count; } catch { }

            Paragraphs paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;
            int index = 0;
            for (int i = 1; i <= count; i++)
            {
                Paragraph p = null;
                try { p = paragraphs[i]; }
                catch { continue; }
                if (p == null) continue;

                index++;

                if (restrictTo != null && !IsInRestriction(p, restrictTo))
                {
                    result.Features.Add(new ParagraphFeatures { Index = index, IsEmpty = true, Text = "" });
                    continue;
                }

                ParagraphFeatures f = ExtractFeatures(p, index);
                result.Features.Add(f);
            }

            watch.Stop();
            result.Duration = watch.Elapsed;
            return result;
        }

        private static bool IsInRestriction(Paragraph p, Range restrict)
        {
            try
            {
                return p.Range.End <= restrict.End && p.Range.Start >= restrict.Start;
            }
            catch
            {
                return false;
            }
        }

        private ParagraphFeatures ExtractFeatures(Paragraph p, int index)
        {
            var f = new ParagraphFeatures { Index = index };

            Range range = p.Range;

            // ---- 文本 ----
            string text = "";
            try { text = range.Text ?? ""; } catch { }
            text = text.TrimEnd('\r', '\a', '\v');
            f.Text = text;
            f.IsEmpty = text.Trim().Length == 0;

            // ---- 字体特征（首字符 run）----
            try
            {
                Font font = range.Font;
                f.FontSize = WordFormatUtils.NormalizeSize(font.Size);
                f.BoldRaw = font.Bold;                 // -1/0/9999999
                f.ColorRgb = WordFormatUtils.WdColorToRgb((int)font.Color);
            }
            catch { }

            // ---- 段落格式 ----
            try
            {
                int align = Convert.ToInt32(p.Alignment);
                f.Alignment = AlignmentKey(align);
            }
            catch { f.Alignment = "?"; }

            // ---- 结构 ----
            try { f.InTable = Convert.ToBoolean(range.get_Information(WdInformation.wdWithInTable)); }
            catch { f.InTable = false; }
            if (f.InTable)
            {
                try { f.IsFirstTableRow = Convert.ToInt32(range.Cells[1].RowIndex) == 1; }
                catch { f.IsFirstTableRow = false; }
            }

            try
            {
                int listType = Convert.ToInt32(range.ListFormat.ListType);
                f.HasAutoNumbering = listType != (int)WdListType.wdListNoNumbering;
            }
            catch { }

            try
            {
                int inlineShapes = range.InlineShapes.Count;
                f.HasInlineImage = inlineShapes > 0;
            }
            catch { }

            try
            {
                int fieldCount = range.Fields.Count;
                f.HasFields = fieldCount > 0;
                if (fieldCount > 0)
                {
                    // 目录（TOC）等大域段落按域代码段处理
                    for (int fi = 1; fi <= fieldCount && fi <= 3; fi++)
                    {
                        int type;
                        try { type = Convert.ToInt32(range.Fields[fi].Type); }
                        catch { break; }
                        if (type == (int)WdFieldType.wdFieldTOC ||
                            type == (int)WdFieldType.wdFieldIndex ||
                            type == (int)WdFieldType.wdFieldRef)
                        {
                            f.IsTocField = true;
                            break;
                        }
                    }
                }
            }
            catch { }

            // 已有 Heading 样式级别：优先 OutlineLevel（对本地化样式名免疫）
            try
            {
                int outline = Convert.ToInt32(p.OutlineLevel);
                if (outline >= (int)WdOutlineLevel.wdOutlineLevel1 && outline <= (int)WdOutlineLevel.wdOutlineLevel3)
                    f.HeadingStyleLevel = outline - (int)WdOutlineLevel.wdOutlineLevel1 + 1;
            }
            catch { }

            // ---- 模式特征（纯文本规则）----
            f.NumberingPattern = DetectNumberingPattern(text);
            f.CaptionPattern = PatternCaption.IsMatch(text);

            var math = MathSegmenter.Analyze(text);
            f.MathDensity = math.Density;
            f.HasEquals = math.HasEquals;
            f.StandaloneMath = math.Standalone;

            return f;
        }

        public static int DetectNumberingPattern(string text)
        {
            if (string.IsNullOrEmpty(text)) return RuleTable.PatternNone;
            if (PatternChapter.IsMatch(text)) return RuleTable.PatternChapter;
            if (PatternNumeric3.IsMatch(text)) return RuleTable.PatternNumeric3;
            if (PatternNumeric2.IsMatch(text)) return RuleTable.PatternNumeric2;
            if (PatternChinese.IsMatch(text)) return RuleTable.PatternChinese;
            if (PatternCircled.IsMatch(text)) return RuleTable.PatternCircled;
            if (PatternBracket.IsMatch(text)) return RuleTable.PatternBracket;
            return RuleTable.PatternNone;
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
