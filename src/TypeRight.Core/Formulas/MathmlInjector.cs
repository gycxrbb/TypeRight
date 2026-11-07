using System;
using System.Collections.Generic;
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
        /// <summary>分段计时累计（性能归因用，FormulaConverter 读取并清零；毫秒）。</summary>
        public static long MsScratch, MsCopy;

        private static long Now() { return System.Diagnostics.Stopwatch.GetTimestamp(); }
        private static long Delta(long start) { return (Now() - start) * 1000 / System.Diagnostics.Stopwatch.Frequency; }

        /// <summary>文档级批量注入：全部公式合并进一个临时 docx（每公式一段），草稿 InsertFile 一次，
        /// 降序逐个 FormattedText 赋值（降序保证未处理目标位置不漂移）。
        /// 返回逐项赋值成败；**不做逐项公式区校验**（254 处 × 2 次 COM 读是纯开销）——
        /// 静默失败由调用方以文档级公式区计数终检兜底（O(1)），计数不符才逐项定位。</summary>
        public static bool[] InjectAll(Document doc, Range[] targets, string[] mathmls, string xslPath, Document scratch)
        {
            int n = targets.Length;
            var ok = new bool[n];
            if (doc == null || scratch == null || n == 0 || n != mathmls.Length) return ok;
            if (string.IsNullOrEmpty(xslPath) || !File.Exists(xslPath)) return ok;

            string tempFile = null;
            try
            {
                // 1) 全部 MathML → OMML
                var ommls = new string[n];
                for (int i = 0; i < n; i++)
                {
                    string omml = XslMathmlToOmml(mathmls[i], xslPath);
                    int s = omml.IndexOf("<m:oMath", StringComparison.Ordinal);
                    if (s < 0) return ok;
                    ommls[i] = omml.Substring(s).Trim();
                }

                // 2) 单个 docx 承载全部公式
                tempFile = Path.Combine(Path.GetTempPath(), "typeright-math-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".docx");
                File.WriteAllBytes(tempFile, BuildMinimalDocx(ommls));

                // 3) 草稿一次装载
                long t0 = Now();
                if (!LoadParagraphsIntoScratch(scratch, tempFile, n))
                {
                    MsScratch += Delta(t0);
                    return ok;
                }
                MsScratch += Delta(t0);

                // 草稿段枚举一次（Paragraphs[i] 索引器是 O(n) 重扫，逐个索引 254 段会平方级）
                var paras = new List<NetOffice.WordApi.Paragraph>(n);
                foreach (NetOffice.WordApi.Paragraph sp in WordCompat.ParagraphsOf(scratch))
                {
                    paras.Add(sp);
                    if (paras.Count >= n) break;
                }
                if (paras.Count < n) return ok;

                // 4) 降序逐个复制（目标位置稳定）
                for (int i = n - 1; i >= 0; i--)
                {
                    try
                    {
                        long t1 = Now();
                        Range srcPara = paras[i].Range;
                        Range src = scratch.Range(srcPara.Start, srcPara.End - 1);
                        targets[i].FormattedText = src.FormattedText;
                        MsCopy += Delta(t1);
                        ok[i] = true;   // 赋值成功即记；公式区是否真建立由调用方终检兜底
                    }
                    catch { }
                }
                return ok;
            }
            catch
            {
                return ok;
            }
            finally
            {
                try { if (tempFile != null && File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }
        }

        /// <summary>清空草稿 → InsertFile → 规范到恰好 expectedParas 段（多余尾段删除）。</summary>
        private static bool LoadParagraphsIntoScratch(Document scratch, string tempFile, int expectedParas)
        {
            try
            {
                scratch.Content.Delete();
                scratch.Range(0, 0).InsertFile(tempFile);
                int guard = 0;
                while (scratch.Paragraphs.Count > expectedParas && guard++ < 10)
                {
                    Range tail = scratch.Range(scratch.Content.End - 1, scratch.Content.End);
                    if ((tail.Text ?? "") != "\r") break;
                    tail.Delete();
                }
                return scratch.Paragraphs.Count == expectedParas;
            }
            catch { return false; }
        }
        /// <summary>注入成功返回 true；任何失败返回 false（调用方回落 Functions 路径）。
        /// segStart/segLen/paraLen 用于判断片段在段中位置，InsertFile 插入完整段落后按需合并拆分。
        /// scratch：复用的隐藏草稿文档（批量转换时开一次、全程复用）；
        /// null = 自开自关（旧行为，兼容单次调用）。频闪根因：每个公式 Documents.Open 一次
        /// 临时文档会在用户窗口顶部队列闪现（实测 254 处公式 = 254 次开合 + ~150ms/个）。</summary>
        public static bool TryInject(Document doc, Range target, string mathml, string xslPath,
            int segStart = -1, int segLen = 0, int paraLen = -1, Document scratch = null)
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

                // 3) 内联注入：草稿文档承载公式 → FormattedText 赋值到目标区间（不引入段落边界）
                bool ownScratch = scratch == null;
                try
                {
                    long t0 = Now();
                    if (ownScratch)
                    {
                        scratch = doc.Application.Documents.Open(tempFile);
                    }
                    else if (!LoadIntoScratch(scratch, tempFile))
                    {
                        MsScratch += Delta(t0);
                        return false;
                    }
                    Range src = scratch.Range(0, scratch.Content.End - 1);
                    MsScratch += Delta(t0);

                    long t1 = Now();
                    target.FormattedText = src.FormattedText;
                    MsCopy += Delta(t1);
                }
                finally
                {
                    if (ownScratch) { try { scratch.Close(false); } catch { } }
                }

                // 4) 成功判据
                try
                {
                    if (target.OMaths.Count == 0)
                    {
                        var probe = target.Duplicate;
                        probe.Collapse(NetOffice.WordApi.Enums.WdCollapseDirection.wdCollapseEnd);
                        probe.MoveEnd(NetOffice.WordApi.Enums.WdUnits.wdCharacter, 1);
                        if (probe.OMaths.Count == 0) return false;
                    }
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

        /// <summary>把临时 docx 内容装进复用草稿：清空 → InsertFile → 删掉插入带来的多余尾段，
        /// 保证草稿只剩"公式¶"一段（Range(0, End-1) 即纯公式内容，与自开路径等价）。</summary>
        private static bool LoadIntoScratch(Document scratch, string tempFile)
        {
            return LoadParagraphsIntoScratch(scratch, tempFile, 1);
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

        /// <summary>把 m:oMath 元素串包进最小合法 docx（zip：ContentTypes + rels + document.xml）。
        /// 单元素 = 单段；多元素 = 每元素一段（批量注入用）。</summary>
        private static byte[] BuildMinimalDocx(params string[] mathElements)
        {
            const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            var body = new StringBuilder();
            foreach (string el in mathElements)
                body.Append("<w:p>").Append(el).Append("</w:p>");

            string documentXml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:document xmlns:w=\"" + w + "\" xmlns:m=\"" + m + "\" xmlns:r=\"" + r + "\">" +
                "<w:body>" + body + "</w:body></w:document>";

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
