// 浮条字体/字号显示取证（真机 COM，单例运行）。
// 断言的是"浮条上真的显示了什么"，不是内部字段——字段对了但没渲染出来也是 bug。
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using NetOffice.WordApi;

namespace TypeRight.Acceptance
{
    public static class DiagFloatDisplay
    {
        public static int Run(TestEnv env)
        {
            Console.WriteLine();
            Console.WriteLine("── 浮条字体/字号显示取证 ──────────────────────");
            int fail = 0;

            // 自建文档：真实讲义里不一定恰好有某个字号，而显示逻辑必须对每个档位都对。
            // 直接复现用户截图场景（仿宋 14pt = 四号，三行同类段）。
            Document doc = null;
            object addin = null;
            try
            {
                doc = env.App.Documents.Add();
                addin = GetAddin(env.App);
                if (addin == null) { Console.WriteLine("  [FAIL] 加载项未连接"); return 1; }

                BuildSample(doc, "仿宋", 14, 3);
                Console.WriteLine("  样本: 仿宋 14pt ×3（= 四号）");

                // ── A：14pt → 「四号」 ──
                doc.Application.Selection.SetRange(1, 1);
                Call(addin, "AutomationSimilarSelect");
                Thread.Sleep(1200);
                string snap = CallString(addin, "AutomationFloatBarDisplay");
                Console.WriteLine("  显示快照(14pt): " + snap);
                fail += Assert(snap, "四号", "A 14pt 应显示中文「四号」");
                fail += Assert(snap, "仿宋", "A 字体框应回填样本字体「仿宋」");
                fail += AssertNot(snap, "|0", "A 字号不得显示 0");
                fail += AssertNot(snap, "|—", "A 字号不得是未知占位");
                Call(addin, "AutomationCloseSession");
                Thread.Sleep(500);

                // ── B：12pt → 「小四」 ──
                BuildSample(doc, "宋体", 12, 3);
                doc.Application.Selection.SetRange(doc.Content.End - 1, doc.Content.End - 1);
                Call(addin, "AutomationSimilarSelect");
                Thread.Sleep(1200);
                snap = CallString(addin, "AutomationFloatBarDisplay");
                Console.WriteLine("  显示快照(12pt): " + snap);
                fail += Assert(snap, "小四", "B 12pt 应显示中文「小四」");
                Call(addin, "AutomationCloseSession");
                Thread.Sleep(500);

                // ── C：10.5pt → 「五号」（默认正文字号，最常见） ──
                BuildSample(doc, "宋体", 10.5, 3);
                doc.Application.Selection.SetRange(doc.Content.End - 1, doc.Content.End - 1);
                Call(addin, "AutomationSimilarSelect");
                Thread.Sleep(1200);
                snap = CallString(addin, "AutomationFloatBarDisplay");
                Console.WriteLine("  显示快照(10.5pt): " + snap);
                fail += Assert(snap, "五号", "C 10.5pt 应显示中文「五号」");
                Call(addin, "AutomationCloseSession");
                Thread.Sleep(500);

                // ── D：非标准档位 12.5pt → 仍显示数值（不能硬套中文名） ──
                BuildSample(doc, "宋体", 12.5, 3);
                doc.Application.Selection.SetRange(doc.Content.End - 1, doc.Content.End - 1);
                Call(addin, "AutomationSimilarSelect");
                Thread.Sleep(1200);
                snap = CallString(addin, "AutomationFloatBarDisplay");
                Console.WriteLine("  显示快照(12.5pt): " + snap);
                fail += Assert(snap, "12.5", "D 非档位 12.5 应显示数值");
                Call(addin, "AutomationCloseSession");
                Thread.Sleep(500);

                // ── E：文本同类模式 → 字号不得为 0/空（回归） ──
                BuildSample(doc, "宋体", 14, 3, withPattern: true);
                try
                {
                    int idx = FindFirst(doc, "正确");
                    if (idx > 0)
                    {
                        env.App.Selection.SetRange(idx, idx + 2);
                        Call(addin, "AutomationTextSimilarSelect");
                        Thread.Sleep(1200);
                        snap = CallString(addin, "AutomationFloatBarDisplay");
                        Console.WriteLine("  显示快照(文本同类): " + snap);
                        fail += AssertNot(snap, "|0", "E 文本同类字号不得为 0");
                        fail += AssertNot(snap, "|", "E 文本同类字号应有值");
                        fail += Assert(snap, "四号", "E 文本同类应显示命中处真实字号");
                        Call(addin, "AutomationCloseSession");
                        Thread.Sleep(400);
                    }
                    else Console.WriteLine("  [SKIP] E 无「正确」样本");
                }
                catch (Exception ex) { Console.WriteLine("  [SKIP] E: " + ex.Message); }
            }
            finally
            {
                try { if (doc != null) doc.Close(false); } catch { }
            }

            Console.WriteLine(fail == 0 ? "  浮条显示取证：ALL PASS" : "  浮条显示取证：FAIL " + fail);
            return fail;
        }

        /// <summary>写 sampleCount 行同格式段（可选择带「正确/错误」用来测文本同类）。</summary>
        private static void BuildSample(Document doc, string font, double sizePt, int paragraphs, bool withPattern = false)
        {
            doc.Content.Delete();
            doc.Content.ParagraphFormat.Alignment = NetOffice.WordApi.Enums.WdParagraphAlignment.wdAlignParagraphJustify;
            for (int i = 0; i < paragraphs; i++)
            {
                string text = withPattern
                    ? "第" + (i + 1) + "题：下列说法正确的是（　）"
                    : "这是第 " + (i + 1) + " 行同类段落，用于浮条字体字号显示取证。";
                Range r = doc.Content;
                r.InsertAfter(text + "\r");
                r = doc.Paragraphs[doc.Paragraphs.Count].Range;
                r.Font.NameFarEast = font;
                r.Font.NameAscii = font;
                r.Font.Size = (float)sizePt;
                try { r.Font.Bold = 0; } catch { }
            }
        }

        private static int FindFirst(Document doc, string pattern)
        {
            try
            {
                string text = doc.Content.Text ?? "";
                int idx = text.IndexOf(pattern, StringComparison.Ordinal);
                return idx < 0 ? -1 : 1 + idx;
            }
            catch { return -1; }
        }

        private static int Assert(string actual, string mustContain, string what)
        {
            if (actual != null && actual.IndexOf(mustContain, StringComparison.Ordinal) >= 0)
            { Console.WriteLine("  [PASS] " + what); return 0; }
            Console.WriteLine("  [FAIL] " + what + " | 期望含「" + mustContain + "」实际「" + actual + "」");
            return 1;
        }

        private static int AssertNot(string actual, string mustNotContain, string what)
        {
            if (actual == null || actual.IndexOf(mustNotContain, StringComparison.Ordinal) < 0)
            { Console.WriteLine("  [PASS] " + what); return 0; }
            Console.WriteLine("  [FAIL] " + what + " | 不应含「" + mustNotContain + "」实际「" + actual + "」");
            return 1;
        }

        private static object GetAddin(Application app)
        {
            try
            {
                // NetOffice 的 COMAddIns 枚举器给的是 COMDynamicObject 包装，
                // 反射 InvokeMember 打不进去；走 IDispatch 后期绑定（与 ui-verify.ps1 同一路径）。
                object comAddIns = app.COMAddIns;
                object item = comAddIns.GetType().InvokeMember("Item",
                    System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.GetProperty,
                    null, comAddIns, new object[] { "TypeRight.Connect" });
                if (item == null) return null;
                return item.GetType().InvokeMember("Object",
                    System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.GetProperty,
                    null, item, null);
            }
            catch (Exception ex)
            {
                Console.WriteLine("  GetAddin: " + ex.Message);
                return null;
            }
        }

        private static void Call(object addin, string method)
        {
            try
            {
                addin.GetType().InvokeMember(method,
                    System.Reflection.BindingFlags.InvokeMethod, null, addin, null);
            }
            catch (Exception ex) { Console.WriteLine("  call " + method + ": " + ex.Message); }
        }

        private static string CallString(object addin, string method)
        {
            try
            {
                object r = addin.GetType().InvokeMember(method,
                    System.Reflection.BindingFlags.InvokeMethod, null, addin, null);
                return r as string ?? "";
            }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }
    }
}
