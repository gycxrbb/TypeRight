using System;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>引导句感知验证：★ 框段（#3）光标在引导句/正文/均匀段三种位置分别匹配。</summary>
    internal static class DiagP7
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var swAll = System.Diagnostics.Stopwatch.StartNew();
            Action<string> step = m =>
            {
                Console.WriteLine("[" + swAll.ElapsedMilliseconds + "ms] " + m);
                Console.Out.Flush();
            };
            step("open doc");
            Document doc = env.App.Documents.Open(env.OriginalDocx, false, true);
            step("doc opened");
            try
            {
                Paragraph box = doc.Paragraphs[3];
                step("para3 got");
                var boxSig = FormatSignatureReader.Read(box);
                step("sig read");
                int boundary = SimilarSelector.DetectLeadBoundary(box, boxSig);
                step("lead boundary = " + boundary);
                Console.WriteLine("★ 框段 #3: lead边界=" + boundary + "（段起点 " + box.Range.Start + "）");
                if (boundary > 0)
                {
                    string leadText = (doc.Range(box.Range.Start, boundary).Text ?? "").TrimEnd('\r', '\a');
                    step("引导句 [" + (leadText.Length > 30 ? leadText.Substring(0, 30) + "…" : leadText) + "]");
                }

                Test(doc, box, 3, "光标在引导句内（offset=3）", step);
                Test(doc, box, 150, "光标在正文内（offset=150）", step);

                // 均匀正文段（宋体正文）
                foreach (Paragraph p in doc.Paragraphs)
                {
                    var s = FormatSignatureReader.Read(p);
                    if (s.IsEmpty || s.EastAsiaFont == null || !s.EastAsiaFont.Contains("宋体")) continue;
                    if (s.SizePt < 10.4 || s.SizePt > 10.6) continue;
                    Test(doc, p, 5, "光标在均匀正文段", step);
                    break;
                }
                step("done");
            }
            finally { try { doc.Close(false); } catch { } }
        }

        private static void Test(Document doc, Paragraph sample, int offset, string label, Action<string> step)
        {
            step("match begin: " + label);
            var r = SimilarSelector.MatchUnits(doc, sample, offset, FormatSignature.IgnoreDims.None);
            step("match end: " + label);
            int leads = 0, bodies = 0;
            string firstText = "";
            foreach (var u in r.Units)
            {
                if (u.IsLead) leads++; else bodies++;
                if (firstText.Length == 0)
                {
                    try
                    {
                        string t = (doc.Range(u.Start, Math.Min(u.End, u.Start + 24)).Text ?? "").TrimEnd('\r', '\a');
                        firstText = "[" + t + "]";
                    }
                    catch { }
                }
            }
            step("== " + label + " → 命中 " + r.Count + " 处（引导句 " + leads + " / 正文·整段 " + bodies +
                "）样本单元=" + (r.SampleIsLead ? "引导句" : "正文") + " 首个命中 " + firstText);
        }
    }
}
