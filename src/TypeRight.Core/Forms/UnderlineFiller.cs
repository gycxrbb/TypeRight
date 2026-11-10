using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Forms
{
    public sealed class UnderlineFillResult
    {
        public bool Success;
        public int BalancedCount;
        public string Message;
    }

    /// <summary>
    /// 能力四：表头下划线零位移填空（Underline Blank Guard）。
    /// 针对 `学院：________  学号：________` 场景：
    /// ① 若光标位于已打字但把右侧顶歪的下划线区间（如 `____经济学院____` 或带下划线的 `   经济学院   `），
    ///    自动吞掉对应宽度的多余空格/`_`，将右侧“学号”拉回原位；
    /// ② 若光标位于空白下划线处，支持传入/弹窗输入填空文字，测量原空白精确磅值宽度后定宽回填（字多自动缩字），
    ///    保证右侧文字零位移。
    /// </summary>
    public static class UnderlineFiller
    {
        private const float MinFontSizePt = 7.5f;

        /// <summary>
        /// 智能执行下划线填空或修复当前段落中的所有下划线填空位。
        /// promptInput: 当光标处于纯空白下划线时，用于向用户索取填写文本的回调（返回 null 表示取消）。
        /// </summary>
        public static UnderlineFillResult FillOrBalance(Application app, Func<string, string> promptInput = null)
        {
            var result = new UnderlineFillResult();
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
                result.Message = "请先打开文档并将光标放在表头下划线处。";
                return result;
            }

            FormGeometry.EnsurePrintView(app);

            Paragraph para = null;
            try { if (sel.Paragraphs != null && sel.Paragraphs.Count > 0) para = sel.Paragraphs[1]; }
            catch { }
            if (para == null)
            {
                result.Message = "请先将光标点进含下划线的表头段落（如「学院：______ 学号：______」）。";
                return result;
            }

            int cursorPos = sel.Range.Start;
            List<BlankSpan> spans = DetectBlankSpansInParagraph(doc, para);
            if (spans.Count == 0)
            {
                result.Message = "当前段落未检测到下划线填空区（支持空格+下划线或连续 _ 字符）。";
                return result;
            }

            // 优先定位光标所在的下划线区间
            BlankSpan targetSpan = null;
            foreach (BlankSpan sp in spans)
            {
                if (cursorPos >= sp.StartDocPos - 1 && cursorPos <= sp.EndDocPos + 1)
                {
                    targetSpan = sp;
                    break;
                }
            }

            // 情况 A：光标在某个纯空白下划线区间内 → 弹出轻量输入框让用户填词，定宽零位移写入
            if (targetSpan != null && string.IsNullOrWhiteSpace(targetSpan.InnerUserText))
            {
                string labelHint = ExtractPrecedingLabel(doc, para.Range.Start, targetSpan.StartDocPos);
                string userText = promptInput != null ? promptInput(labelHint) : null;
                if (string.IsNullOrEmpty(userText))
                {
                    result.Message = "";
                    return result;
                }

                UndoRecord undo = WordCompat.BeginUndo(app, "TypeRight 下划线零位移填空");
                try
                {
                    FillSpanToExactWidth(doc, targetSpan, userText.Trim());
                    result.Success = true;
                    result.BalancedCount = 1;
                    result.Message = "✓ 已定宽填入「" + userText.Trim() + "」，右侧文字零位移（Ctrl+Z 可撤销）";
                    return result;
                }
                catch (Exception ex)
                {
                    result.Message = "下划线填空失败：" + ex.Message;
                    return result;
                }
                finally
                {
                    WordCompat.EndUndo(undo);
                }
            }

            // 情况 B：用户已经直接在下划线中间打了字（两侧残留多余空格或 `_` 把后面顶歪了）→ 一键吞空格配平全段
            UndoRecord undoBalance = WordCompat.BeginUndo(app, "TypeRight 下划线对齐修复");
            try
            {
                int fixedCount = 0;
                // 从右向左处理，避免字符增删导致后续区间偏移
                for (int i = spans.Count - 1; i >= 0; i--)
                {
                    BlankSpan sp = spans[i];
                    if (!string.IsNullOrWhiteSpace(sp.InnerUserText) && (sp.LeadingHalfUnits + sp.TrailingHalfUnits) >= 2)
                    {
                        if (AbsorbExcessPaddingInSpan(doc, sp))
                            fixedCount++;
                    }
                }

                if (fixedCount == 0 && targetSpan == null && spans.Count > 0)
                {
                    // 当前段全是空白下划线且光标不在其中一处，则对第一处空白下划线触发填空
                    BlankSpan firstBlank = spans[0];
                    string labelHint = ExtractPrecedingLabel(doc, para.Range.Start, firstBlank.StartDocPos);
                    string userText = promptInput != null ? promptInput(labelHint) : null;
                    if (!string.IsNullOrEmpty(userText))
                    {
                        FillSpanToExactWidth(doc, firstBlank, userText.Trim());
                        result.Success = true;
                        result.BalancedCount = 1;
                        result.Message = "✓ 已定宽填入「" + userText.Trim() + "」，右侧文字零位移（Ctrl+Z 可撤销）";
                        return result;
                    }
                }

                if (fixedCount > 0)
                {
                    result.Success = true;
                    result.BalancedCount = fixedCount;
                    result.Message = "✓ 已自动吞掉 " + fixedCount + " 处下划线多余占位空格，恢复右侧对齐（Ctrl+Z 可撤销）";
                }
                else
                {
                    result.Message = "请将光标点在空白下划线上填空，或在已打字的下划线处点击自动对齐。";
                }
                return result;
            }
            catch (Exception ex)
            {
                result.Message = "下划线修复失败：" + ex.Message;
                return result;
            }
            finally
            {
                WordCompat.EndUndo(undoBalance);
            }
        }

        private sealed class BlankSpan
        {
            public int StartDocPos;
            public int EndDocPos;
            public string RawText;
            public string InnerUserText;
            public int LeadingHalfUnits;
            public int TrailingHalfUnits;
        }

        /// <summary>
        /// 扫描段落内的下划线填空区间：
        /// 支持 ① 带下划线格式的空格/文字区间；② 由连续 `_` / `＿` 构成的区间（含中间已打字如 `___信息学院___`）。
        /// </summary>
        private static List<BlankSpan> DetectBlankSpansInParagraph(Document doc, Paragraph para)
        {
            var list = new List<BlankSpan>();
            try
            {
                int pStart = para.Range.Start;
                int pEnd = para.Range.End - 1; // 排除段落标记 \r
                if (pEnd <= pStart) return list;

                string text = doc.Range(pStart, pEnd).Text ?? "";
                if (text.Length == 0) return list;

                int n = text.Length;
                bool[] isUnderlineSlot = new bool[n];

                for (int i = 0; i < n; i++)
                {
                    char c = text[i];
                    if (c == '_' || c == '＿')
                    {
                        isUnderlineSlot[i] = true;
                        continue;
                    }

                    // 检查字符是否开启了 Word 下划线格式
                    try
                    {
                        Range chRng = doc.Range(pStart + i, pStart + i + 1);
                        WdUnderline u = chRng.Font.Underline;
                        if (u != WdUnderline.wdUnderlineNone && (int)u != 9999999)
                        {
                            isUnderlineSlot[i] = true;
                        }
                    }
                    catch { }
                }

                // 同时识别：夹在连续 `_` 左右两侧之间的已填文字（如 `___计算机学院___`，中间的字可能没下划线）
                var underscoreSandwich = new Regex(@"([_＿]{2,})([^_＿\r\n：:]{1,20})([_＿]{2,})");
                foreach (Match m in underscoreSandwich.Matches(text))
                {
                    for (int k = m.Index; k < m.Index + m.Length && k < n; k++)
                        isUnderlineSlot[k] = true;
                }

                int idx = 0;
                while (idx < n)
                {
                    if (!isUnderlineSlot[idx]) { idx++; continue; }
                    int runStart = idx;
                    while (idx < n && isUnderlineSlot[idx]) idx++;
                    int runEnd = idx; // exclusive

                    string sub = text.Substring(runStart, runEnd - runStart);
                    ParseSpanContent(sub, out string innerText, out int leadHalf, out int trailHalf);

                    // 至少包含 2 个半角单位的空白/`_`，才算是下划线填空位
                    if (leadHalf + trailHalf >= 2)
                    {
                        list.Add(new BlankSpan
                        {
                            StartDocPos = pStart + runStart,
                            EndDocPos = pStart + runEnd,
                            RawText = sub,
                            InnerUserText = innerText,
                            LeadingHalfUnits = leadHalf,
                            TrailingHalfUnits = trailHalf
                        });
                    }
                }
            }
            catch { }
            return list;
        }

        private static void ParseSpanContent(string sub, out string innerText, out int leadHalf, out int trailHalf)
        {
            leadHalf = 0;
            trailHalf = 0;
            int i = 0;
            while (i < sub.Length && IsPadChar(sub[i]))
            {
                leadHalf += CharHalfWidth(sub[i]);
                i++;
            }
            int j = sub.Length - 1;
            while (j >= i && IsPadChar(sub[j]))
            {
                trailHalf += CharHalfWidth(sub[j]);
                j--;
            }
            innerText = (i <= j) ? sub.Substring(i, j - i + 1) : "";
        }

        private static bool IsPadChar(char c)
        {
            return c == ' ' || c == '\u3000' || c == '_' || c == '＿' || c == '\u00A0';
        }

        public static int CharHalfWidth(char c)
        {
            if (c == '\u3000' || c == '＿') return 2;
            return c > 127 ? 2 : 1;
        }

        public static int TextHalfWidth(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            int total = 0;
            foreach (char c in s) total += CharHalfWidth(c);
            return total;
        }

        /// <summary>
        /// 定宽零位移写入：先测原空白区间的精确磅值宽度，写入文字并动态配平左右半角空格 + 字距/字号，
        /// 保证填写后区间右端 X 坐标与填写前完全一致。
        /// </summary>
        private static void FillSpanToExactWidth(Document doc, BlankSpan span, string userText)
        {
            Range spanRng = doc.Range(span.StartDocPos, span.EndDocPos);
            Range startPtRng = doc.Range(span.StartDocPos, span.StartDocPos);
            Range endPtRng = doc.Range(span.EndDocPos, span.EndDocPos);

            float startX = FormGeometry.GetPageX(startPtRng);
            float origEndX = FormGeometry.GetPageX(endPtRng);
            float targetWidthPt = (startX >= 0f && origEndX > startX + 4f) ? (origEndX - startX) : -1f;

            float origFontSize = 10.5f;
            try { if (spanRng.Font.Size > 5f && spanRng.Font.Size < 72f) origFontSize = spanRng.Font.Size; } catch { }

            int origHalfUnits = span.LeadingHalfUnits + span.TrailingHalfUnits;
            int textHalfUnits = TextHalfWidth(userText);
            int remainHalfUnits = Math.Max(0, origHalfUnits - textHalfUnits);
            int leftSpaces = remainHalfUnits / 2;
            int rightSpaces = remainHalfUnits - leftSpaces;

            string composed = new string(' ', leftSpaces) + userText + new string(' ', rightSpaces);
            spanRng.Text = composed;
            spanRng.Font.Underline = WdUnderline.wdUnderlineSingle;

            // 若可精确测量页面 X 坐标，则微调空格数、字符间距或字号，使右端坐标严格对齐 origEndX
            if (targetWidthPt > 4f)
            {
                BalanceRangeWidthToTarget(doc, spanRng, userText, leftSpaces, rightSpaces, origEndX, origFontSize);
            }
        }

        /// <summary>
        /// 当用户已经直接在下划线中间打了字（如原本 10 个空格，打了 4 个汉字变成 `     经济学院     ` 共 18 单位），
        /// 自动吞掉与文字等宽的 8 个半角空格（同时将野生 `_` 规范化为真下划线空格），使右侧文字回到原位。
        /// </summary>
        private static bool AbsorbExcessPaddingInSpan(Document doc, BlankSpan span)
        {
            int currentPadUnits = span.LeadingHalfUnits + span.TrailingHalfUnits;
            int textUnits = TextHalfWidth(span.InnerUserText);

            // 用户打字前原本的空白单位数 = 当前剩余左右空白单位数（因为用户在空白中间直接插入了文字，原空白一格未删）
            // 所以应当从当前左右空白中扣除 textUnits 个半角单位！
            int newPadUnits = Math.Max(0, currentPadUnits - textUnits);
            int leftSpaces = newPadUnits / 2;
            int rightSpaces = newPadUnits - leftSpaces;

            Range spanRng = doc.Range(span.StartDocPos, span.EndDocPos);
            float origFontSize = 10.5f;
            try { if (spanRng.Font.Size > 5f && spanRng.Font.Size < 72f) origFontSize = spanRng.Font.Size; } catch { }

            string composed = new string(' ', leftSpaces) + span.InnerUserText + new string(' ', rightSpaces);
            spanRng.Text = composed;
            spanRng.Font.Underline = WdUnderline.wdUnderlineSingle;

            // 若填写的字本身比原空白还长（textUnits > currentPadUnits），按比例微缩字距/字号以抵消溢出宽度
            if (textUnits > currentPadUnits && currentPadUnits >= 4)
            {
                float ratio = (float)currentPadUnits / textUnits;
                float newSize = Math.Max(MinFontSizePt, (float)Math.Round(origFontSize * ratio * 2f) / 2f);
                try
                {
                    spanRng.Font.Spacing = -0.5f;
                    if (newSize < origFontSize)
                        spanRng.Font.Size = newSize;
                }
                catch { }
            }
            return true;
        }

        private static void BalanceRangeWidthToTarget(
            Document doc, Range spanRng, string userText,
            int initLeftSpaces, int initRightSpaces, float targetEndX, float origFontSize)
        {
            try
            {
                int left = initLeftSpaces;
                int right = initRightSpaces;

                // 最多微调 6 步半角空格，消除中西文字符宽度换算误差
                for (int step = 0; step < 6; step++)
                {
                    Range endCheck = doc.Range(spanRng.End, spanRng.End);
                    float curEndX = FormGeometry.GetPageX(endCheck);
                    if (curEndX < 0f) break;

                    float diff = curEndX - targetEndX;
                    if (Math.Abs(diff) <= 2.2f) break;

                    if (diff > 2.2f && (left + right) > 0)
                    {
                        // 偏宽：减少 1 个半角空格
                        if (right >= left && right > 0) right--;
                        else if (left > 0) left--;
                        spanRng.Text = new string(' ', left) + userText + new string(' ', right);
                        spanRng.Font.Underline = WdUnderline.wdUnderlineSingle;
                    }
                    else if (diff < -2.2f)
                    {
                        // 偏窄：补充 1 个半角空格
                        if (left <= right) left++;
                        else right++;
                        spanRng.Text = new string(' ', left) + userText + new string(' ', right);
                        spanRng.Font.Underline = WdUnderline.wdUnderlineSingle;
                    }
                    else break;
                }

                // 若左右空格已减到 0 但文字仍然超出原下划线宽度，则启动微缩字距 + 0.5pt 阶梯降字号
                Range finalEnd = doc.Range(spanRng.End, spanRng.End);
                float finalX = FormGeometry.GetPageX(finalEnd);
                if (finalX > targetEndX + 2.5f && left == 0 && right == 0)
                {
                    try { spanRng.Font.Spacing = -0.5f; } catch { }
                    for (float sz = origFontSize - 0.5f; sz >= MinFontSizePt; sz -= 0.5f)
                    {
                        finalEnd = doc.Range(spanRng.End, spanRng.End);
                        if (FormGeometry.GetPageX(finalEnd) <= targetEndX + 2.0f) break;
                        try { spanRng.Font.Size = sz; } catch { break; }
                    }
                }
            }
            catch { }
        }

        private static string ExtractPrecedingLabel(Document doc, int paraStart, int spanStart)
        {
            try
            {
                int len = Math.Min(12, spanStart - paraStart);
                if (len <= 0) return "";
                string pre = doc.Range(spanStart - len, spanStart).Text ?? "";
                pre = pre.TrimEnd(' ', '\t', '\u3000', '：', ':');
                int lastSpace = pre.LastIndexOfAny(new[] { ' ', '\t', '\u3000' });
                if (lastSpace >= 0 && lastSpace + 1 < pre.Length)
                    pre = pre.Substring(lastSpace + 1);
                return pre.Trim();
            }
            catch { return ""; }
        }
    }
}
