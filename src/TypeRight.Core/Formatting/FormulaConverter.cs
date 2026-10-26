using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Xsl;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formulas;
using TypeRight.Profiles;
using TypeRight.Scanning;

namespace TypeRight.Formatting
{
    /// <summary>
    /// 纯文本公式 → OMath（PRD §6）：片段切分 → UnicodeMath 线性串 → OMaths.Add → BuildUp。
    /// 失败片段原样保留为文本并计入统计，绝不留半转换残骸。
    /// </summary>
    public sealed class FormulaConverter
    {
        private readonly Document _doc;
        private readonly FormulaOptions _options;
        private readonly string _mml2OmmlPath;

        public FormulaConverter(Document doc, FormulaOptions options, string mml2OmmlXslPath)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            _doc = doc;
            _options = options ?? new FormulaOptions();
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

        /// <summary>
        /// 转换一个段落内的公式片段。wholeLine=true 时（formulaLine 角色）转换整段。
        /// 已是公式对象的区间（OMaths）自动跳过，保证幂等。
        /// </summary>
        public void ConvertParagraph(ParagraphFeatures feature, Paragraph paragraph, ApplyStats stats, bool wholeLine)
        {
            if (feature == null || paragraph == null || stats == null) return;
            if (feature.IsEmpty || feature.HasFields) return;

            Range baseRange = paragraph.Range;
            string liveText;
            try { liveText = baseRange.Text ?? ""; }
            catch { return; }

            List<MathSegmenter.Segment> segments;
            if (wholeLine)
            {
                string trimmed = liveText.TrimEnd('\r', '\a', '\v');
                int lead = liveText.Length - liveText.TrimStart(' ').Length;
                if (MathSegmenter.IsFormulaCandidate(trimmed.Trim()))
                    segments = new List<MathSegmenter.Segment>
                    {
                        new MathSegmenter.Segment { Start = lead, Length = trimmed.Trim().Length, Text = trimmed.Trim() }
                    };
                else
                    segments = MathSegmenter.FindSegments(trimmed);
            }
            else
            {
                segments = MathSegmenter.FindSegments(liveText.TrimEnd('\r', '\a', '\v'));
            }
            if (segments.Count == 0) return;

            // 降序处理，先改后面的片段不影响前面的偏移
            for (int i = segments.Count - 1; i >= 0; i--)
            {
                var seg = segments[i];
                Range rng = null;
                try
                {
                    rng = baseRange.Duplicate;
                    rng.Start = baseRange.Start + seg.Start;
                    rng.End = rng.Start + seg.Length;
                }
                catch { continue; }

                // 已是公式区间则跳过（幂等）
                try { if (rng.OMaths.Count > 0) continue; } catch { }

                string original = seg.Text;
                if (TryConvertRange(rng, original, stats)) stats.FormulasConverted++;
            }
        }

        /// <summary>单个片段的转换。失败时尽力还原原文并返回 false。</summary>
        private bool TryConvertRange(Range rng, string originalText, ApplyStats stats)
        {
            string mapped = UnicodeMathMapper.Map(originalText);
            int mathCountBefore = 0;
            try { mathCountBefore = rng.OMaths.Count; } catch { }

            try
            {
                if (!string.Equals(mapped, originalText, StringComparison.Ordinal))
                    rng.Text = mapped;

                // NetOffice：OMaths.Add(range) 返回覆盖新建公式对象的 Range
                Range mathRange = _doc.OMaths.Add(rng);
                if (mathRange == null) throw new InvalidOperationException("OMaths.Add 返回空。");
                OMath om = mathRange.OMaths[1];
                if (om == null) throw new InvalidOperationException("未取到公式对象。");
                om.BuildUp();

                // 成功性验证：区间已进入公式集合
                int mathCountAfter;
                try { mathCountAfter = rng.OMaths.Count; } catch { mathCountAfter = 1; }
                if (mathCountAfter <= mathCountBefore) throw new InvalidOperationException("公式对象未建立。");

                if (_options.SizePt.HasValue)
                {
                    try { om.Range.Font.Size = (float)_options.SizePt.Value; } catch { }
                }
                return true;
            }
            catch
            {
                Restore(rng, originalText, mathCountBefore);
                stats.FormulasFailed++;
                return false;
            }
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
                // 还原失败是极端情况（不应发生）：至少不让异常冒泡打断整体排版
            }
        }

        // ---- MML2OMML.XSL 兜底管线（PRD §6.2.1）----
        private static XslCompiledTransform _mmlTransform;
        private static string _mmlTransformPath;

        /// <summary>MathML → OMML（Office 官方 XSLT）。加载失败抛异常，由调用方决定是否走该兜底。</summary>
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
            using (var reader = new System.IO.StringReader(mathml))
            using (var writer = new System.IO.StringWriter())
            {
                using (var xmlReader = System.Xml.XmlReader.Create(reader))
                {
                    _mmlTransform.Transform(xmlReader, new System.Xml.Xsl.XsltArgumentList(), writer);
                }
                return writer.ToString();
            }
        }

        /// <summary>把线性文本包装为 mtext 的 MathML（兜底管线的降级输入：至少成为 Cambria Math 公式对象）。</summary>
        public static string WrapLinearAsMtextMathml(string linearText)
        {
            string escaped = (linearText ?? "")
                .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
            return "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mtext>" + escaped + "</mtext></math>";
        }
    }
}
