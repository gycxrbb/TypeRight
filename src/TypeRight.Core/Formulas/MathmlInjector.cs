using System;
using System.Diagnostics;
using NetOffice.WordApi;

namespace TypeRight.Formulas
{
    /// <summary>
    /// MathML → Word 公式注入（TC-19 管线）：MathML 文本放剪贴板 UnicodeText →
    /// 选中原片段 → Selection.Paste，Word 自动将 MathML 转为 OMML（探针 P4 实证）。
    /// </summary>
    public static class MathmlInjector
    {
        /// <summary>注入成功返回 true；任何失败返回 false（调用方回落 Functions 路径）。</summary>
        public static bool TryInject(Application app, Document doc, Range target, string mathml)
        {
            if (app == null || doc == null || target == null || string.IsNullOrEmpty(mathml)) return false;

            int zonesBefore = 0;
            try { zonesBefore = doc.OMaths.Count; } catch { }

            string prevClip = null;
            bool clipRestorable = false;
            try
            {
                try
                {
                    if (System.Windows.Forms.Clipboard.ContainsText())
                    {
                        prevClip = System.Windows.Forms.Clipboard.GetText();
                        clipRestorable = true;
                    }
                }
                catch { }

                System.Windows.Forms.Clipboard.SetText(mathml, System.Windows.Forms.TextDataFormat.UnicodeText);
                target.Select();
                app.Selection.Paste();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (clipRestorable)
                {
                    try { System.Windows.Forms.Clipboard.SetText(prevClip, System.Windows.Forms.TextDataFormat.UnicodeText); } catch { }
                }
            }

            // 成功判据（探针实证）：
            // A 按纯文本贴入（段落出现明文 <math）→ Undo + 失败
            // B/C 目标位置已成为公式区（新建或与相邻区合并）→ 成功
            // D 什么都没贴（目标仍是纯文本）→ Undo + 失败
            try
            {
                string paraText = target.Paragraphs[1].Range.Text ?? "";
                if (paraText.Contains("<math"))
                {
                    try { doc.Undo(); } catch { }
                    return false;
                }
            }
            catch { }

            try
            {
                if (target.OMaths.Count > 0) return true;
                // 目标区间可能已被合并区整体吞并导致区间折叠：以粘贴点后一字符复核
                var probe = target.Duplicate;
                probe.Collapse(NetOffice.WordApi.Enums.WdCollapseDirection.wdCollapseEnd);
                probe.MoveEnd(NetOffice.WordApi.Enums.WdUnits.wdCharacter, 1);
                if (probe.OMaths.Count > 0) return true;
            }
            catch { }

            try { doc.Undo(); } catch { }
            return false;
        }
    }
}
