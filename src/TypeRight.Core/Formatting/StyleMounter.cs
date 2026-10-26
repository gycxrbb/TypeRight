using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Profiles;
using TypeRight.Roles;

namespace TypeRight.Formatting
{
    /// <summary>
    /// 样式挂载（PRD §3 规则 / §5.1）：wordStyle 存在时，先挂样式、再把字体属性写入样式定义（全文联动）。
    /// </summary>
    public sealed class StyleMounter
    {
        /// <summary>定位（或创建）角色对应的样式。内置样式走 WdBuiltinStyle 枚举（对界面语言免疫），自定义名按名查找/创建。</summary>
        public Style EnsureStyle(Document doc, string wordStyle)
        {
            if (string.IsNullOrEmpty(wordStyle)) return null;

            var builtin = WordFormatUtils.TryGetBuiltinStyle(wordStyle);
            if (builtin.HasValue)
            {
                try { return doc.Styles[(int)builtin.Value]; }
                catch { return null; }
            }

            // 自定义样式：按名查找；找不到则创建段落型样式
            try { return doc.Styles[wordStyle]; }
            catch
            {
                try
                {
                    doc.Styles.Add(wordStyle, WdStyleType.wdStyleTypeParagraph);
                    return doc.Styles[wordStyle];
                }
                catch { return null; }
            }
        }

        /// <summary>把角色格式写入样式定义（字体族分中西文通道写入，互不污染）。</summary>
        public bool WriteStyleDefinition(Document doc, Style style, RoleFormat format, ApplyStats stats)
        {
            if (style == null || format == null) return false;
            bool changed = false;

            Font font = null;
            try { font = style.Font; } catch { }
            if (font != null)
            {
                if (!string.IsNullOrEmpty(format.EastAsiaFont))
                    changed |= WordFormatUtils.SetIfChanged(() => font.NameFarEast, v => font.NameFarEast = v, format.EastAsiaFont);
                if (!string.IsNullOrEmpty(format.LatinFont))
                {
                    changed |= WordFormatUtils.SetIfChanged(() => font.NameAscii, v => font.NameAscii = v, format.LatinFont);
                    changed |= WordFormatUtils.SetIfChanged(() => font.NameOther, v => font.NameOther = v, format.LatinFont);
                }
                if (format.SizePt.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => font.Size, v => font.Size = v, (float)format.SizePt.Value, WordFormatUtils.FloatEq);
                if (format.Bold.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => font.Bold, v => font.Bold = v, format.Bold.Value ? -1 : 0, WordFormatUtils.IntEq);
                if (!string.IsNullOrEmpty(format.Color))
                {
                    int? rgb = WordFormatUtils.ParseHexColor(format.Color);
                    if (rgb.HasValue)
                        changed |= WordFormatUtils.SetIfChanged(() => font.Color, v => font.Color = v,
                            (NetOffice.WordApi.Enums.WdColor)WordFormatUtils.RgbToWdColor(rgb.Value), (a, b) => a == b);
                }
            }

            ParagraphFormat pf = null;
            try { pf = style.ParagraphFormat; } catch { }
            if (pf != null)
            {
                if (format.AlignmentValue != ProfileAlignment.None)
                    changed |= WordFormatUtils.SetIfChanged(() => pf.Alignment, v => pf.Alignment = v,
                        WordFormatUtils.ToWdAlignment(format.AlignmentValue), (a, b) => a == b);
                if (format.LineSpacing.HasValue)
                {
                    float before;
                    try { before = pf.LineSpacing; } catch { before = -1; }
                    int ruleBefore;
                    try { ruleBefore = Convert.ToInt32(pf.LineSpacingRule); } catch { ruleBefore = -1; }
                    WordFormatUtils.ApplyLineSpacing(pf, format.LineSpacing.Value);
                    try
                    {
                        if (before == pf.LineSpacing && ruleBefore == Convert.ToInt32(pf.LineSpacingRule)) { /* unchanged */ }
                        else changed = true;
                    }
                    catch { changed = true; }
                }
                if (format.SpaceBeforePt.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => pf.SpaceBefore, v => pf.SpaceBefore = v, (float)format.SpaceBeforePt.Value, WordFormatUtils.FloatEq);
                if (format.SpaceAfterPt.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => pf.SpaceAfter, v => pf.SpaceAfter = v, (float)format.SpaceAfterPt.Value, WordFormatUtils.FloatEq);
                if (format.FirstLineIndentPt.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => pf.FirstLineIndent, v => pf.FirstLineIndent = v, (float)format.FirstLineIndentPt.Value, WordFormatUtils.FloatEq);
                if (format.OutlineLevel.HasValue)
                    changed |= WordFormatUtils.SetIfChanged(() => pf.OutlineLevel, v => pf.OutlineLevel = v,
                        WordFormatUtils.ToWdOutlineLevel(format.OutlineLevel.Value), (a, b) => a == b);
            }

            if (changed && stats != null) stats.StyleDefinitionsUpdated++;
            return changed;
        }

        /// <summary>段落挂载样式。</summary>
        public bool MountParagraph(Paragraph paragraph, Style style)
        {
            if (paragraph == null || style == null) return false;
            try
            {
                object current = paragraph.Style;
                // NetOffice 返回 Style 代理或名称字符串
                string currentName = current is Style ? ((Style)current).NameLocal : Convert.ToString(current);
                if (string.Equals(currentName, style.NameLocal, StringComparison.Ordinal)) return false;
                paragraph.Style = style;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    /// <summary>一次排版执行的改动统计（PRD §5.1 末尾）。</summary>
    public sealed class ApplyStats
    {
        public int ChangedParagraphs;
        public int StyleDefinitionsUpdated;
        public int FormulasConverted;
        public int FormulasFailed;
        public readonly System.Collections.Generic.Dictionary<string, int> ParagraphsByRole =
            new System.Collections.Generic.Dictionary<string, int>();

        public int Skipped;

        public void CountRole(ParagraphRole role)
        {
            string key = role.ToLabel();
            int v;
            ParagraphsByRole.TryGetValue(key, out v);
            ParagraphsByRole[key] = v + 1;
        }

        public string Summary()
        {
            var parts = new System.Collections.Generic.List<string>();
            foreach (var kv in ParagraphsByRole) parts.Add(kv.Key + " " + kv.Value + " 段");
            parts.Add("样式定义更新 " + StyleDefinitionsUpdated + " 个");
            parts.Add("公式转换 " + FormulasConverted + " 处" + (FormulasFailed > 0 ? "（失败保留原文 " + FormulasFailed + " 处）" : ""));
            parts.Add("跳过 " + Skipped + " 段");
            return string.Join(" ｜ ", parts);
        }
    }
}
