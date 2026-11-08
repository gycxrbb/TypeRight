using System;
using System.Collections.Generic;
using NetOffice.WordApi;

namespace TypeRight.Signature
{
    /// <summary>字义匹配（2026-11 出卷场景：「正确/错误」着重号批量标注）。
    /// 与格式签名匹配正交：不看格式只看字符，全文找同一文本的出现处。
    /// 默认排除否定前缀（不/非/未——"不正确"里的"正确"不标，语义翻转防护，可关）；
    /// 公式区（OMML）内不匹配（那是公式对象不是文本）。</summary>
    public static class TextMatcher
    {
        private static readonly char[] NegationPrefixes = { '不', '非', '未' };

        /// <summary>全文找 pattern 的全部出现（非重叠，文档序）。
        /// excludeNegation=true 时，出现处前一个字符是否定前缀则跳过。</summary>
        public static List<MatchUnit> FindOccurrences(Document doc, string pattern, bool excludeNegation)
        {
            var units = new List<MatchUnit>();
            if (doc == null || string.IsNullOrEmpty(pattern)) return units;
            List<int[]> zones = SimilarSelector.MathZoneRanges(doc);
            foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
            {
                int pStart, pEnd;
                string text;
                try
                {
                    Range r = p.Range;
                    pStart = r.Start; pEnd = r.End;
                    text = r.Text ?? "";
                }
                catch { continue; }
                int idx = 0;
                while (idx + pattern.Length <= text.Length)
                {
                    int hit = text.IndexOf(pattern, idx, StringComparison.Ordinal);
                    if (hit < 0) break;
                    idx = hit + pattern.Length;                       // 非重叠
                    int abs = pStart + hit, absEnd = abs + pattern.Length;
                    if (absEnd > pEnd) break;
                    if (excludeNegation && hit > 0 && Array.IndexOf(NegationPrefixes, text[hit - 1]) >= 0)
                        continue;                                     // "不/非/未正确"里的"正确"不标
                    if (Overlaps(zones, abs, absEnd))
                        continue;                                     // 公式区内不匹配
                    units.Add(new MatchUnit { Paragraph = p, Start = abs, End = absEnd });
                }
            }
            return units;
        }

        private static bool Overlaps(List<int[]> ranges, int s, int e)
        {
            if (ranges == null) return false;
            foreach (int[] r in ranges)
                if (s < r[1] && e > r[0]) return true;
            return false;
        }
    }
}
