using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using NetOffice.WordApi;

namespace TypeRight.Formulas
{
    /// <summary>
    /// MathML → OMML → 临时 docx → Range.InsertFile 注入（TC-19 管线，PRD §6.2.1）。
    /// 不占用剪贴板（全局用户资源，禁止污染——剪贴板方案已废弃）。
    /// </summary>
    public static class MathmlInjector
    {
        /// <summary>注入成功返回 true；任何失败返回 false（调用方回落 Functions 路径）。
        /// segStart/segLen/paraLen 用于判断片段在段中位置，InsertFile 插入完整段落后按需合并拆分。</summary>
        public static bool TryInject(Document doc, Range target, string mathml, string xslPath,
            int segStart = -1, int segLen = 0, int paraLen = -1)
        {
            if (doc == null || target == null || string.IsNullOrEmpty(mathml)) return false;
            if (string.IsNullOrEmpty(xslPath) || !File.Exists(xslPath)) return false;

            string tempFile = null;
            try
            {
                // 1) MathML → OMML（Office 自带 XSL）
                string omml = XslMathmlToOmml(mathml, xslPath);
                int i = omml.IndexOf("<m:oMath", StringComparison.Ordinal);
                if (i < 0) return false;
                string mathElement = omml.Substring(i).Trim();

                // 2) OMML → 最小 docx
                tempFile = Path.Combine(Path.GetTempPath(), "typeright-math-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".docx");
                File.WriteAllBytes(tempFile, BuildMinimalDocx(mathElement));

                // 3) 内联注入：临时文档承载公式 → FormattedText 赋值到目标区间（不引入段落边界）
                NetOffice.WordApi.Document scratch = null;
                try
                {
                    scratch = doc.Application.Documents.Open(tempFile);
                    Range src = scratch.Range(0, scratch.Content.End - 1);
                    target.FormattedText = src.FormattedText;
                }
                finally
                {
                    try { if (scratch != null) { scratch.Close(false); } } catch { }
                }

                // 4) 成功判据 + 合并段落拆分
                try
                {
                    if (target.OMaths.Count == 0)
                    {
                        var probe = target.Duplicate;
                        probe.Collapse(NetOffice.WordApi.Enums.WdCollapseDirection.wdCollapseEnd);
                        probe.MoveEnd(NetOffice.WordApi.Enums.WdUnits.wdCharacter, 1);
                        if (probe.OMaths.Count == 0) return false;
                    }

                    // FormattedText 不拆段，无需合并
                    return true;
                }
                catch { return false; }
            }
            catch
            {
                return false;
            }
            finally
            {
                try { if (tempFile != null && File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }
        }

        /// <summary>
        /// InsertFile 插入完整段落（自带段落标记），会把原段落一分为三：前文¶公式¶后文。
        /// 按片段原位置删除多余的段落标记，恢复单段结构。
        /// </summary>
        private static void MergeSplitParagraphs(Document doc, Range target, int segStart, int segLen, int paraLen)
        {
            try
            {
                OMath om = target.OMaths[1];
                if (om == null) return;
                Paragraph mathPara = om.Range.Paragraphs[1];
                if (mathPara == null) return;

                bool known = segStart >= 0 && paraLen >= 0 && segLen > 0;
                bool hasTextBefore = !known || segStart > 0;
                bool hasTextAfter = known && segStart + segLen < paraLen;

                if (hasTextBefore)
                {
                    int s = mathPara.Range.Start;
                    if (s > 0)
                    {
                        Range del = doc.Range(s - 1, s);
                        if ((del.Text ?? "") == "\r") del.Delete();
                    }
                }

                if (hasTextAfter)
                {
                    int e = mathPara.Range.End - 1;
                    if (e > mathPara.Range.Start)
                    {
                        Range del2 = doc.Range(e, e + 1);
                        if ((del2.Text ?? "") == "\r") del2.Delete();
                    }
                }
            }
            catch { }
        }


        private static System.Xml.Xsl.XslCompiledTransform _xsl;
        private static string _xslPath;

        public static string TransformForProbe(string mathml, string xslPath) { return XslMathmlToOmml(mathml, xslPath); }

        private static string XslMathmlToOmml(string mathml, string xslPath)
        {
            if (_xsl == null || _xslPath != xslPath)
            {
                var t = new System.Xml.Xsl.XslCompiledTransform();
                t.Load(xslPath);
                _xsl = t;
                _xslPath = xslPath;
            }
            using (var reader = new StringReader(mathml))
            using (var writer = new StringWriter())
            {
                using (var xmlReader = System.Xml.XmlReader.Create(reader))
                {
                    _xsl.Transform(xmlReader, null, writer);
                }
                return writer.ToString();
            }
        }

        /// <summary>把 m:oMath 元素包进最小合法 docx（zip：ContentTypes + rels + document.xml）。</summary>
        private static byte[] BuildMinimalDocx(string mathElement)
        {
            const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            string documentXml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:document xmlns:w=\"" + w + "\" xmlns:m=\"" + m + "\" xmlns:r=\"" + r + "\">" +
                "<w:body><w:p>" + mathElement + "</w:p></w:body></w:document>";

            const string contentTypes =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "</Types>";

            const string rels =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                "</Relationships>";

            using (var ms = new MemoryStream())
            {
                using (var zip = new ZipArchive(ms, ZipArchiveMode.Create))
                {
                    WriteEntry(zip, "[Content_Types].xml", contentTypes);
                    WriteEntry(zip, "_rels/.rels", rels);
                    WriteEntry(zip, "word/document.xml", documentXml);
                }
                return ms.ToArray();
            }
        }

        private static void WriteEntry(ZipArchive zip, string path, string content)
        {
            ZipArchiveEntry entry = zip.CreateEntry(path, CompressionLevel.Fastest);
            using (var sw = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
            {
                sw.Write(content);
            }
        }
    }
}
