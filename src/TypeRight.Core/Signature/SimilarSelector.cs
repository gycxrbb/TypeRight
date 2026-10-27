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
    /// 性能目标：1000 段 &lt; 1s。
    /// </summary>
    public static class SimilarSelector
    {
        /// <summary>样本签名 → 全文匹配。</summary>
        public static SimilarMatchResult Match(Document doc, FormatSignature sample, FormatSignature.IgnoreDims ignore)
        {
            var watch = Stopwatch.StartNew();
            var result = new SimilarMatchResult { SampleSignature = sample, Ignore = ignore };
            if (doc == null || sample.IsEmpty) return result;

            Paragraphs paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;
            for (int i = 1; i <= count; i++)
            {
                Paragraph p = null;
                try { p = paragraphs[i]; }
                catch { continue; }
                if (p == null) continue;

                FormatSignature s = FormatSignatureReader.Read(p);
                if (s.IsEmpty) continue;          // 空段/无可见文本不参与
                if (s.Matches(sample, ignore)) result.Paragraphs.Add(p);
            }

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
