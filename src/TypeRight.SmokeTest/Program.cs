using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Formatting;
using TypeRight.Profiles;
using TypeRight.Signature;

namespace TypeRight.SmokeTest
{
    /// <summary>
    /// PRD v2 黄金旅程端到端（真实 Word）：
    /// ① 同类选择（严格/宽松命中数）→ 浮动条语义批量改字体（中西文联动）
    /// ② 每步独立撤销 ③ 公式转换计数 ④ 存方案 → 新副本一键回放 → 结果一致 + 幂等。
    /// </summary>
    internal static class Program
    {
        private static int _failures;

        private static void Check(bool condition, string description)
        {
            Console.WriteLine((condition ? "  [PASS] " : "  [FAIL] ") + description);
            if (!condition) _failures++;
        }

        private static void Header(string title)
        {
            Console.WriteLine();
            Console.WriteLine("=== " + title + " ===");
        }

        private static int Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;
                Console.WriteLine("TypeRight v2 冒烟测试（真实 Word · 指认式批量排版）");
                return Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("[FATAL] " + ex);
                return 2;
            }
        }

        private static int Run()
        {
            Application app = new Application();
            app.Visible = false;
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            Document doc = null;
            try
            {
                doc = SampleDocBuilder.Build(app);

                // ── 1. 签名与同类选择 ──
                Header("1. 同类选择（签名匹配）");
                Paragraph title = FindParagraph(doc, "一元二次方程讲义");
                Paragraph module1 = FindParagraph(doc, "模块一");
                Paragraph module2 = FindParagraph(doc, "模块二");
                Paragraph section11 = FindParagraph(doc, "1.1 ");
                Paragraph body = FindParagraph(doc, "对于方程");

                var titleSig = FormatSignatureReader.Read(title);
                Console.WriteLine("  标题签名: " + titleSig);

                var strict = SimilarSelector.Match(doc, titleSig, FormatSignature.IgnoreDims.None);
                Check(strict.Count == 1, "严格模式：主标题只命中自己（实际 " + strict.Count + "）");

                var looseSize = SimilarSelector.Match(doc, titleSig, FormatSignature.IgnoreDims.SizePt);
                Check(looseSize.Count == 3, "忽略字号：主标题 + 2 个模块标题 = 3 段（实际 " + looseSize.Count + "）");

                var sectionSig = FormatSignatureReader.Read(section11);
                var sections = SimilarSelector.Match(doc, sectionSig, FormatSignature.IgnoreDims.None);
                Check(sections.Count == 3, "小节标题严格命中 3 段（1.1 / 1.1.2 / 1.2）（实际 " + sections.Count + "）");

                var bodySig = FormatSignatureReader.Read(body);
                var bodies = SimilarSelector.Match(doc, bodySig, FormatSignature.IgnoreDims.Bold | FormatSignature.IgnoreDims.Color);
                Console.WriteLine("  正文宽松命中: " + bodies.Count + " 段");
                Check(bodies.Count >= 4, "正文宽松（忽略加粗/颜色）命中 ≥ 4 段（实际 " + bodies.Count + "）");

                // 代理管理开关对比（性能实验）
                var swPM = System.Diagnostics.Stopwatch.StartNew();
                app.Factory.Settings.EnableProxyManagement = false;
                var pmMatch = SimilarSelector.Match(doc, bodySig, FormatSignature.IgnoreDims.Bold | FormatSignature.IgnoreDims.Color);
                swPM.Stop();
                app.Factory.Settings.EnableProxyManagement = true;
                Console.WriteLine("  [实验] 关闭代理管理: " + pmMatch.Count + " 段, " + swPM.ElapsedMilliseconds + "ms/全文档");
                // 性能：本测试经 DCOM（进程外 Word，每次 COM 调用 ~1ms 编组），
                // 每段 ~14ms 属编组开销主导；真实插件在 Word 进程内执行（vtable 直调，预期 <1s/千段）。
                // 进程内实测以 addin-log 的「同类匹配: N 段 / Xms」为准（首次点击同类选择即可看到）。
                Perf1000(doc, bodySig);

                // ── 2. 批量格式 + 中西文联动 + 独立撤销 ──
                Header("2. 批量格式与撤销粒度");
                string fontBefore = title.Range.Font.NameFarEast;
                var stats1 = BatchFormatter.Apply(app, doc, looseSize.Paragraphs,
                    new FormatChanges { EastAsiaFont = "楷体", LatinFont = "Times New Roman" }, "同类格式");
                Console.WriteLine("  批量改楷体：目标 " + looseSize.Count + " 段，改动 " + stats1.Changed);
                Check(stats1.Changed >= 3, "3 段发生改动");
                Check(title.Range.Font.NameFarEast == "楷体", "主标题中文 = 楷体");
                Check(title.Range.Font.NameAscii == "Times New Roman", "主标题西文联动 = Times New Roman");
                Check(title.Range.Font.Size == 22, "字号未被触碰（仍是 22）");

                // 每步独立撤销：一次 Ctrl+Z 只退这一步
                var stats2 = BatchFormatter.Apply(app, doc, looseSize.Paragraphs,
                    new FormatChanges { EastAsiaFont = "黑体" }, "同类格式2");
                Check(stats2.Changed == 3, "第二步改黑体 3 段");
                bool undone = doc.Undo();
                Check(undone && title.Range.Font.NameFarEast == "楷体",
                    "单次撤销只退第二步：中文回到楷体（实际 " + title.Range.Font.NameFarEast + "）");
                undone = doc.Undo();
                Check(undone && title.Range.Font.NameFarEast == fontBefore,
                    "再退一次回到原始字体（实际 " + title.Range.Font.NameFarEast + "）");

                // 局部加粗保留（正文段里的加粗词在批量改字体后仍在）
                var boldWord = FindRange(doc, "判别式与求根公式的应用");
                Check(boldWord != null && boldWord.Font.Bold == -1, "段内局部加粗不受批量字体影响");

                // ── 3. 公式转换 ──
                Header("3. 公式转换");
                var convStats = ConvertAll(app, doc);
                Console.WriteLine("  " + convStats.SummaryText());
                Check(convStats.Converted >= 5, "公式转换 ≥ 5 处（实际 " + convStats.Converted + "）");
                Check(convStats.Failed == 0, "零失败（实际 " + convStats.Failed + "）");
                Check(doc.OMaths.Count == convStats.Converted, "公式对象数与计数一致（" + doc.OMaths.Count + "）");
                bool undoFormula = doc.Undo();
                Check(undoFormula && doc.OMaths.Count == 0, "单次撤销公式转换（残留 " + doc.OMaths.Count + "）");

                // ── 4. 方案：录制 → 回放 ──
                Header("4. 方案录制与回放");
                var profile = new ProfileV2
                {
                    Id = "smoke-" + DateTime.Now.ToString("HHmmss"),
                    Name = "冒烟方案",
                    ConvertFormulas = true,
                    Rules = new List<ProfileRule>
                    {
                        new ProfileRule { Order = 1, Match = ToMatch(titleSig, FormatSignature.IgnoreDims.SizePt),
                                          Apply = new ApplySpec { EastAsiaFont = "楷体", LatinFont = "Times New Roman" } },
                        new ProfileRule { Order = 2, Match = ToMatch(sectionSig, FormatSignature.IgnoreDims.None),
                                          Apply = new ApplySpec { EastAsiaFont = "黑体", LatinFont = "Times New Roman" } },
                        new ProfileRule { Order = 3, Match = new MatchSpec { CatchAll = true },
                                          Apply = new ApplySpec { EastAsiaFont = "宋体", LatinFont = "Times New Roman" } }
                    }
                };
                Check(ProfileV2Validator.Validate(profile).Count == 0, "方案校验通过");

                string applied1 = ApplyProfileReport(app, doc, profile);
                Console.WriteLine("  首次应用: " + applied1);
                Check(title.Range.Font.NameFarEast == "楷体" && module1.Range.Font.NameFarEast == "楷体",
                    "回放后主标题/模块 = 楷体");
                Check(section11.Range.Font.NameFarEast == "黑体", "回放后小节 = 黑体");
                Check(body.Range.Font.NameFarEast == "宋体", "回放后正文（兜底）= 宋体");
                Check(doc.OMaths.Count > 0, "回放包含公式转换");

                // 幂等：第二次 0 改动
                var second = BatchFormatter.Apply(app, doc, new List<Paragraph> { title },
                    new FormatChanges { EastAsiaFont = "楷体" }, null);
                Check(second.Changed == 0, "幂等：重复应用相同字体 0 改动");
                var convStats2 = ConvertAll(app, doc);
                Check(convStats2.Converted == 0, "幂等：公式二次转换 0 处");

                // 回放到全新副本结果一致
                Document fresh = SampleDocBuilder.Build(app);
                string applied2 = ApplyProfileReport(app, fresh, profile);
                Check(fresh.Paragraphs.Count == doc.Paragraphs.Count, "副本与原文档段落一致");
                Check(FindParagraph(fresh, "一元二次方程讲义").Range.Font.NameFarEast == "楷体", "新副本主标题 = 楷体");
                fresh.Close(false);

                // 保存输出供人工检查
                string outDir = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "out");
                System.IO.Directory.CreateDirectory(outDir);
                string outPath = System.IO.Path.Combine(outDir, "smoke-v2-result.docx");
                try { doc.SaveAs2(outPath); Console.WriteLine(); Console.WriteLine("结果文档: " + outPath); }
                catch (Exception ex) { Console.WriteLine("保存结果文档失败: " + ex.Message); }
            }
            finally
            {
                try { if (doc != null) { doc.Close(false); doc.Dispose(); } } catch { }
                try { app.Quit(); app.Dispose(); } catch { }
            }

            Console.WriteLine();
            Console.WriteLine(_failures == 0 ? "全部通过 ✔" : "失败 " + _failures + " 项 ✘");
            return _failures == 0 ? 0 : 1;
        }

        private static FormulaStats ConvertAll(Application app, Document doc)
        {
            bool screenOff = false;
            try { app.ScreenUpdating = false; screenOff = true; } catch { }
            var undo = app.UndoRecord;
            undo.StartCustomRecord("公式转换");
            try
            {
                return new FormulaConverter(doc, null, null).ConvertDocument();
            }
            finally
            {
                try { undo.EndCustomRecord(); } catch { }
                if (screenOff) { try { app.ScreenUpdating = true; } catch { } }
            }
        }

        /// <summary>按方案规则逐条执行（与 TypeRightController.ApplyProfile 同语义）。</summary>
        private static string ApplyProfileReport(Application app, Document doc, ProfileV2 profile)
        {
            var report = new List<string>();
            var consumed = new HashSet<int>();
            bool screenOff = false;
            try { app.ScreenUpdating = false; screenOff = true; } catch { }
            var undo = app.UndoRecord;
            undo.StartCustomRecord("应用方案「" + profile.Name + "」");
            try
            {
                foreach (ProfileRule rule in profile.Rules.OrderBy(r => r.Order))
                {
                    var hit = new List<Paragraph>();
                    foreach (Paragraph p in doc.Paragraphs)
                    {
                        try
                        {
                            if (consumed.Contains(p.Range.Start)) continue;
                            var sig = FormatSignatureReader.Read(p);
                            if (!sig.IsEmpty && rule.Match.MatchesPara(sig)) hit.Add(p);
                        }
                        catch { }
                    }
                    foreach (Paragraph p in hit)
                    {
                        try { consumed.Add(p.Range.Start); } catch { }
                    }
                    if (hit.Count > 0)
                    {
                        BatchFormatter.Apply(app, doc, hit, new FormatChanges
                        {
                            EastAsiaFont = rule.Apply.EastAsiaFont,
                            LatinFont = rule.Apply.LatinFont,
                            SizePt = rule.Apply.SizePt,
                            Bold = rule.Apply.Bold
                        }, null, wrapUndo: false);
                    }
                    report.Add("规则" + rule.Order + " 命中 " + hit.Count);
                }
                if (profile.ConvertFormulas)
                {
                    var s = new FormulaConverter(doc, null, null).ConvertDocument();
                    report.Add("公式 " + s.Converted);
                }
            }
            finally
            {
                try { undo.EndCustomRecord(); } catch { }
                if (screenOff) { try { app.ScreenUpdating = true; } catch { } }
            }
            return string.Join(" / ", report);
        }

        private static MatchSpec ToMatch(FormatSignature sig, FormatSignature.IgnoreDims ignore)
        {
            var spec = new MatchSpec
            {
                EastAsiaFont = sig.EastAsiaFont,
                LatinFont = sig.LatinFont,
                Ignore = new List<string>()
            };
            if (!double.IsNaN(sig.SizePt)) spec.SizePt = sig.SizePt;
            if (!sig.BoldMixed) spec.Bold = sig.Bold;
            if (sig.ColorRgb >= 0) spec.Color = "#" + sig.ColorRgb.ToString("X6");
            if ((ignore & FormatSignature.IgnoreDims.SizePt) != 0) spec.Ignore.Add("sizePt");
            if (spec.Ignore.Count == 0) spec.Ignore = null;
            return spec;
        }

        /// <summary>性能：重复匹配直到等效 1000 段规模（进程外 DCOM 预算）。</summary>
        private static void Perf1000(Document doc, FormatSignature sig)
        {
            int paragraphs = doc.Paragraphs.Count;
            int repeats = Math.Max(1, (int)Math.Ceiling(1000.0 / paragraphs));
            var watch = System.Diagnostics.Stopwatch.StartNew();
            for (int i = 0; i < repeats; i++) SimilarSelector.Match(doc, sig, FormatSignature.IgnoreDims.None);
            watch.Stop();
            Console.WriteLine("  性能(DCOM 参考): " + paragraphs + " 段 × " + repeats + " 次 = " + watch.ElapsedMilliseconds + "ms");
        }

        private static Paragraph FindParagraph(Document doc, string startsWith)
        {
            foreach (Paragraph p in doc.Paragraphs)
            {
                try
                {
                    string t = (p.Range.Text ?? "").TrimEnd('\r', '\a');
                    if (t.StartsWith(startsWith, StringComparison.Ordinal)) return p;
                }
                catch { }
            }
            return null;
        }

        private static Range FindRange(Document doc, string text)
        {
            var find = doc.Content;
            find.Find.ClearFormatting();
            find.Find.Text = text;
            find.Find.Forward = true;
            find.Find.Wrap = WdFindWrap.wdFindStop;
            if (find.Find.Execute()) return find;
            return null;
        }
    }
}
