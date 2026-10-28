using System;
using System.Collections.Generic;
using System.Diagnostics;
using NetOffice.WordApi;
using TypeRight.Formatting;
using TypeRight.Profiles;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>PRD §3 黄金旅程（4 轮同类选择改字体 + 1 次公式转换），TC-12/15/16/18 共用。</summary>
    public sealed class JourneyResult
    {
        public int Round1Hits;   // 模块（忽略字号）→ 楷体
        public int Round2Hits;   // 小节（严格）→ 黑体
        public int Round3Hits;   // 正文（忽略加粗+颜色）→ 宋体
        public int Round4Hits;   // 图注（忽略字号+颜色）→ 宋体
        public int CatchAllHits; // 其余段落（表格/★ 等）→ 宋体
        public FormulaStats Formula;
        public double TotalMs;
        public List<OpRecording.Op> Ops = new List<OpRecording.Op>();
    }

    public static class Journey
    {
        public static JourneyResult Run(Application app, Document doc, TestEnv env, bool withFormulas = true)
        {
            var result = new JourneyResult();
            var watch = Stopwatch.StartNew();
            var processed = new HashSet<int>();

            // 第 1 轮：模块标题（忽略字号）→ 楷体
            Paragraph module = env.FindParagraph(doc, "模块一");
            if (module != null)
            {
                var sig = FormatSignatureReader.Read(module);
                var hit = SimilarSelector.Match(doc, module, FormatSignature.IgnoreDims.SizePt);
                result.Round1Hits = hit.Count;
                foreach (Paragraph hitP in hit.Paragraphs) { try { processed.Add(hitP.Range.Start); } catch { } }
                var changes = new FormatChanges { EastAsiaFont = "楷体", LatinFont = "Times New Roman" };
                BatchFormatter.Apply(app, doc, hit.Paragraphs, changes, "同类格式（模块→楷体）");
                result.Ops.Add(new OpRecording.Op { SampleSignature = sig, Ignore = FormatSignature.IgnoreDims.SizePt, Applied = changes, SampleInTable = false });
            }

            // 第 2 轮：小节标题（严格）→ 黑体
            Paragraph section = env.FindParagraph(doc, "1.1 ");
            if (section != null)
            {
                var sig = FormatSignatureReader.Read(section);
                var hit = SimilarSelector.Match(doc, section, FormatSignature.IgnoreDims.None);
                result.Round2Hits = hit.Count;
                foreach (Paragraph hitP in hit.Paragraphs) { try { processed.Add(hitP.Range.Start); } catch { } }
                var changes = new FormatChanges { EastAsiaFont = "黑体", LatinFont = "Times New Roman" };
                BatchFormatter.Apply(app, doc, hit.Paragraphs, changes, "同类格式（小节→黑体）");
                result.Ops.Add(new OpRecording.Op { SampleSignature = sig, Ignore = FormatSignature.IgnoreDims.None, Applied = changes, SampleInTable = false });
            }

            // 第 3 轮：正文系（忽略加粗+颜色）→ 宋体
            Paragraph body = env.FindParagraph(doc, "解一元二次方程的灵魂是");
            if (body != null)
            {
                var sig = FormatSignatureReader.Read(body);
                var hit = SimilarSelector.Match(doc, body, FormatSignature.IgnoreDims.Bold | FormatSignature.IgnoreDims.Color);
                result.Round3Hits = hit.Count;
                foreach (Paragraph hitP in hit.Paragraphs) { try { processed.Add(hitP.Range.Start); } catch { } }
                var changes = new FormatChanges { EastAsiaFont = "宋体", LatinFont = "Times New Roman" };
                BatchFormatter.Apply(app, doc, hit.Paragraphs, changes, "同类格式（正文→宋体）");
                result.Ops.Add(new OpRecording.Op { SampleSignature = sig, Ignore = FormatSignature.IgnoreDims.Bold | FormatSignature.IgnoreDims.Color, Applied = changes, SampleInTable = false });
            }

            // 第 4 轮：图注（忽略字号+颜色）→ 宋体
            Paragraph caption = env.FindParagraph(doc, "【图解例图】");
            if (caption != null)
            {
                var sig = FormatSignatureReader.Read(caption);
                var hit = SimilarSelector.Match(doc, caption, FormatSignature.IgnoreDims.SizePt | FormatSignature.IgnoreDims.Color);
                result.Round4Hits = hit.Count;
                foreach (Paragraph hitP in hit.Paragraphs) { try { processed.Add(hitP.Range.Start); } catch { } }
                var changes = new FormatChanges { EastAsiaFont = "宋体", LatinFont = "Times New Roman" };
                BatchFormatter.Apply(app, doc, hit.Paragraphs, changes, "同类格式（图注→宋体）");
                result.Ops.Add(new OpRecording.Op { SampleSignature = sig, Ignore = FormatSignature.IgnoreDims.SizePt | FormatSignature.IgnoreDims.Color, Applied = changes, SampleInTable = false });
            }

            // 第 5 步（黄金取证）：其余段落（表格内段落、★ 导读等未被前四轮覆盖者）→ 宋体，
            // 对应方案 catchAll 规则；黄金文件中 8 个 ★ 段与全部表格 run 均为 宋体+TNR 且字号/加粗/颜色保源值
            var rest = new List<Paragraph>();
            foreach (Paragraph p in doc.Paragraphs)
            {
                try
                {
                    int st = p.Range.Start;
                    if (processed.Contains(st)) continue;
                    var s2 = FormatSignatureReader.Read(p);
                    if (!s2.IsEmpty) rest.Add(p);
                }
                catch { }
            }
            if (rest.Count > 0)
            {
                result.CatchAllHits = rest.Count;
                var catchAllChanges = new FormatChanges { EastAsiaFont = "宋体", LatinFont = "Times New Roman" };
                BatchFormatter.Apply(app, doc, rest, catchAllChanges, "其余段落→宋体");
                result.Ops.Add(new OpRecording.Op { SampleSignature = new FormatSignature(), Ignore = FormatSignature.IgnoreDims.None, Applied = catchAllChanges, SampleInTable = false });
                result.Ops[result.Ops.Count - 1] = MakeCatchAllOp(catchAllChanges);
            }

            // 第 6 步：公式转换
            if (withFormulas)
            {
                bool screenOff = false;
                try { app.ScreenUpdating = false; screenOff = true; } catch { }
                var undo = app.UndoRecord;
                undo.StartCustomRecord("公式转换");
                try { result.Formula = new FormulaConverter(doc, null, TypeRight.Formatting.FormulaConverter.ResolveMml2OmmlPath(env.App)).ConvertDocument(); }
                finally
                {
                    try { undo.EndCustomRecord(); } catch { }
                    if (screenOff) { try { app.ScreenUpdating = true; } catch { } }
                }
            }

            watch.Stop();
            result.TotalMs = watch.ElapsedMilliseconds;
            return result;
        }

        private static OpRecording.Op MakeCatchAllOp(FormatChanges changes)
        {
            return new OpRecording.Op
            {
                SampleSignature = new FormatSignature(),
                Ignore = FormatSignature.IgnoreDims.None,
                Applied = changes,
                SampleInTable = false,
                IsCatchAll = true
            };
        }

        /// <summary>由旅程操作生成方案（TC-15：签名规则 + catchAll，语义等价结构）。</summary>
        public static ProfileV2 ToProfile(JourneyResult journey)
        {
            return OpRecording.BuildProfile("math-acceptance", "数学讲义", journey.Ops,
                withCatchAll: false, catchAllFont: null, defaultLatinFont: "Times New Roman",
                convertFormulas: journey.Formula != null);
        }
    }
}
