using System;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>复现自动化语境的 RecomputeMatches NRE：可见 Word + 只读 + 第 3 段 + SimilarSelector.Match。</summary>
    internal static class DiagP3
    {
        private static Document _doc;
        private static bool _inEvent;

        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            env.App.Visible = true;
            try { env.App.WindowState = NetOffice.WordApi.Enums.WdWindowState.wdWindowStateMaximize; } catch { }

            _doc = env.App.Documents.Open(env.OriginalDocx, false, true);
            try
            {
                Paragraph sample = _doc.Paragraphs[3];
                sample.Range.Select();
                Console.WriteLine("段3文本: [" + (sample.Range.Text ?? "").Replace("\r", "¶").Replace("\a", "§") + "]");

                var sig = FormatSignatureReader.Read(sample);
                Console.WriteLine("sig = " + sig);

                bool inTable = Convert.ToBoolean(sample.Range.get_Information(
                    NetOffice.WordApi.Enums.WdInformation.wdWithInTable));
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var result = SimilarSelector.Match(_doc, sig, FormatSignature.IgnoreDims.None, inTable);
                Console.WriteLine("[直接调用] Match: " + result.Count + " 段 / " + sw.ElapsedMilliseconds + "ms");

                // 复刻加载项语境：在 WindowSelectionChange 事件处理器内跑同样的 Match
                env.App.WindowSelectionChangeEvent += OnSelChange;
                Console.WriteLine("-- 触发选择变化，事件内跑 Match --");
                _doc.Paragraphs[4].Range.Select();
                System.Threading.Thread.Sleep(2000);
                env.App.WindowSelectionChangeEvent -= OnSelChange;
                Console.WriteLine("done");
            }
            catch (Exception ex)
            {
                Console.WriteLine("EX: " + ex);
            }
            finally
            {
                try { _doc.Close(false); } catch { }
            }
        }

        private static void OnSelChange(NetOffice.WordApi.Selection selection)
        {
            if (_inEvent) return;
            _inEvent = true;
            try
            {
                Paragraph sample = _doc.Paragraphs[3];
                var sig = FormatSignatureReader.Read(sample);
                bool inTable = Convert.ToBoolean(sample.Range.get_Information(
                    NetOffice.WordApi.Enums.WdInformation.wdWithInTable));
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var result = SimilarSelector.Match(_doc, sig, FormatSignature.IgnoreDims.None, inTable);
                Console.WriteLine("[事件内调用] Match: " + result.Count + " 段 / " + sw.ElapsedMilliseconds + "ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine("[事件内] EX: " + ex.GetType().Name + ": " + ex.Message);
            }
            finally { _inEvent = false; }
        }
    }
}
