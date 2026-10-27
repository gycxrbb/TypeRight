using System;
using System.Collections.Generic;
using System.Diagnostics;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formatting;

namespace TypeRight.Signature
{
    /// <summary>同类匹配结果。</summary>
    public sealed class SimilarMatchResult
    {
        /// <summary>命中段落（文档序）。</summary>
        public List<Paragraph> Paragraphs = new List<Paragraph>();
        public FormatSignature SampleSignature;
        public FormatSignature.IgnoreDims Ignore;
        public TimeSpan Duration;

        public int Count { get { return Paragraphs.Count; } }
    }

    /// <summary>
    /// 同类段落匹配（PRD v2 P0-1）：全文段落 + 表格单元格内段落，按签名比对。
    /// 分区规则（黄金文件取证）：正文样本只匹配正文段落，表格样本只匹配表格内段落——
    /// 避免正文样本把表格里的同格式装饰段（如 ★ 导读）一并卷入。
    /// 性能目标：1000 段 &lt; 1s。
    /// </summary>
    public static class SimilarSelector
    {
        /// <summary>样本段落驱动：提取签名 + 判定表格分区后匹配。</summary>
        public static SimilarMatchResult Match(Document doc, Paragraph sampleParagraph, FormatSignature.IgnoreDims ignore)
        {
            bool sampleInTable = false;
            try
            {
                sampleInTable = Convert.ToBoolean(sampleParagraph.Range.get_Information(
                    NetOffice.WordApi.Enums.WdInformation.wdWithInTable));
            }
            catch { }
            return Match(doc, FormatSignatureReader.Read(sampleParagraph), ignore, sampleInTable);
        }

        /// <summary>样本签名 → 全文匹配（分区：sampleInTable 决定命中文本域）。</summary>
        public static SimilarMatchResult Match(Document doc, FormatSignature sample, FormatSignature.IgnoreDims ignore, bool sampleInTable = false)
        {
            var watch = Stopwatch.StartNew();
            var result = new SimilarMatchResult { SampleSignature = sample, Ignore = ignore };
            if (doc == null || sample.IsEmpty) return result;

            Paragraphs paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;
            FormatSignatureReader.TicksRange = FormatSignatureReader.TicksText = FormatSignatureReader.TicksProbe =
                FormatSignatureReader.TicksFont = FormatSignatureReader.TicksAlign = 0;
            long enumStart = Stopwatch.GetTimestamp();
            // 注意：Paragraphs[i] 索引器是 O(n) 重扫（千段文档实测 O(n²)≈8.9s），必须用枚举器顺序遍历
            int index = 0;
            foreach (Paragraph p in paragraphs)
            {
                index++;
                bool inTable = false;
                try { inTable = Convert.ToBoolean(p.Range.get_Information(NetOffice.WordApi.Enums.WdInformation.wdWithInTable)); }
                catch { }
                if (inTable != sampleInTable) continue;   // 正文/表格分区
                FormatSignature s = FormatSignatureReader.Read(p);
                if (s.IsEmpty) continue;          // 空段/无可见文本不参与
                if (s.Matches(sample, ignore)) result.Paragraphs.Add(p);
            }
            count = index;
            FormatSignatureReader.TicksEnumerator = Stopwatch.GetTimestamp() - enumStart
                - FormatSignatureReader.TicksRange - FormatSignatureReader.TicksText
                - FormatSignatureReader.TicksProbe - FormatSignatureReader.TicksFont - FormatSignatureReader.TicksAlign;

            watch.Stop();
            result.Duration = watch.Elapsed;
            return result;
        }

        // ---- 命中呈现：临时高亮（PRD P0-1：不写入正文内容语义，退出即还原）----
        // 只高亮原本“无高亮”的段落，原有高亮的段落保持不动（避免无法精确还原混合高亮）。

        public sealed class HighlightSession
        {
            private readonly Document _doc;
            internal readonly List<int> TouchedStarts = new List<int>();
            internal readonly List<int> TouchedEnds = new List<int>();

            public HighlightSession(Document doc) { _doc = doc; }

            /// <summary>把命中段落标记为黄色（原本无高亮的）。返回本次实际标记的段数。</summary>
            public int Apply(SimilarMatchResult match)
            {
                Clear();
                if (match == null) return 0;
                int applied = 0;
                foreach (Paragraph p in match.Paragraphs)
                {
                    try
                    {
                        Range r = p.Range;
                        WdColorIndex original = r.HighlightColorIndex;
                        if (original == WdColorIndex.wdNoHighlight)
                        {
                            r.HighlightColorIndex = WdColorIndex.wdYellow;
                            TouchedStarts.Add(r.Start);
                            TouchedEnds.Add(r.End);
                            applied++;
                        }
                    }
                    catch { }
                }
                return applied;
            }

            /// <summary>还原本次会话标记的全部高亮。</summary>
            public void Clear()
            {
                for (int i = 0; i < TouchedStarts.Count; i++)
                {
                    try
                    {
                        Range r = _doc.Range(TouchedStarts[i], TouchedEnds[i]);
                        r.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                    }
                    catch { }
                }
                TouchedStarts.Clear();
                TouchedEnds.Clear();
            }
        }
    }
}
