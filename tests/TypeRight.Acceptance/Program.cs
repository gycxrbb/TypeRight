using System;
using System.Collections.Generic;
using System.Text;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Acceptance
{
    /// <summary>
    /// TypeRight 验收测试入口（规格：tests/TypeRight_测试验收规格.md，TC-01~TC-18）。
    /// 用法：TypeRight.Acceptance [用例名过滤，如 TC-01 或 TC-1]
    /// 退出码 = 失败断言数（0 = 全绿）。
    /// </summary>
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            string filter = args.Length > 0 ? args[0] : null;

            Application app = new Application();
            app.Visible = false;
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            var env = new TestEnv(app);

            if (filter == "diag")
            {
                Diag.Run(env);
                try { app.Quit(); app.Dispose(); } catch { }
                return 0;
            }
            if (filter == "mini")
            {
                DiagMini.Run(env);
                try { app.Quit(); app.Dispose(); } catch { }
                return 0;
            }
            if (filter == "xsl")
            {
                DiagXsl.Run(env);
                try { app.Quit(); app.Dispose(); } catch { }
                return 0;
            }
            if (filter == "fb")
            {
                DiagFb.Run(env);
                try { app.Quit(); app.Dispose(); } catch { }
                return 0;
            }
            if (filter == "rpr")
            {
                DiagRpr.Run(env);
                try { app.Quit(); app.Dispose(); } catch { }
                return 0;
            }
            if (filter == "clip")
            {
                DiagClip.Run(env);
                try { app.Quit(); app.Dispose(); } catch { }
                return 0;
            }
            if (filter == "insert")
            {
                DiagInsert.Run(env);
                try { app.Quit(); app.Dispose(); } catch { }
                return 0;
            }
            if (filter == "golden")
            {
                DiagGolden.Run(env);
                try { app.Quit(); app.Dispose(); } catch { }
                return 0;
            }
            if (!System.IO.File.Exists(env.OriginalDocx))
            {
                Console.WriteLine("[FATAL] 找不到测试资产: " + env.OriginalDocx);
                return 2;
            }
            if (!System.IO.File.Exists(env.GoldenDocx))
            {
                Console.WriteLine("[FATAL] 找不到黄金文件: " + env.GoldenDocx);
                return 2;
            }
            Console.WriteLine("原始: " + System.IO.Path.GetFileName(env.OriginalDocx));
            Console.WriteLine("黄金: " + System.IO.Path.GetFileName(env.GoldenDocx));

            var cases = new List<KeyValuePair<string, Action<TestEnv>>>
            {
                new KeyValuePair<string, Action<TestEnv>>("TC-01", CasesA.TC01),
                new KeyValuePair<string, Action<TestEnv>>("TC-02", CasesA.TC02),
                new KeyValuePair<string, Action<TestEnv>>("TC-03", CasesA.TC03),
                new KeyValuePair<string, Action<TestEnv>>("TC-04", CasesA.TC04),
                new KeyValuePair<string, Action<TestEnv>>("TC-05", CasesA.TC05),
                new KeyValuePair<string, Action<TestEnv>>("TC-06", CasesA.TC06),
                new KeyValuePair<string, Action<TestEnv>>("TC-07", CasesA.TC07),
                new KeyValuePair<string, Action<TestEnv>>("TC-08", CasesA.TC08),
                new KeyValuePair<string, Action<TestEnv>>("TC-09", CasesA.TC09),
                new KeyValuePair<string, Action<TestEnv>>("TC-10", CasesB.TC10),
                new KeyValuePair<string, Action<TestEnv>>("TC-11", CasesB.TC11),
                new KeyValuePair<string, Action<TestEnv>>("TC-12", CasesB.TC12),
                new KeyValuePair<string, Action<TestEnv>>("TC-13", CasesB.TC13),
                new KeyValuePair<string, Action<TestEnv>>("TC-14", CasesB.TC14),
                new KeyValuePair<string, Action<TestEnv>>("TC-15", e => { e.RecycleWord(); CasesC.TC15(e); }),
                new KeyValuePair<string, Action<TestEnv>>("TC-16", e => { e.RecycleWord(); CasesC.TC16(e); }),
                new KeyValuePair<string, Action<TestEnv>>("TC-17", e => { e.RecycleWord(); CasesC.TC17(e); }),
                new KeyValuePair<string, Action<TestEnv>>("TC-18", e => { e.RecycleWord(); CasesC.TC18(e); }),
                new KeyValuePair<string, Action<TestEnv>>("TC-19", e => { e.RecycleWord(); CasesD.TC19(e); }),
                new KeyValuePair<string, Action<TestEnv>>("TC-20", e => { e.RecycleWord(); CasesD.TC20(e); }),
                                                            };

            int executed = 0;
            foreach (var kv in cases)
            {
                if (filter != null && !kv.Key.StartsWith(filter, StringComparison.OrdinalIgnoreCase)) continue;
                executed++;
                try
                {
                    kv.Value(env);
                }
                catch (Exception ex)
                {
                    env.Begin(kv.Key);
                    env.Check(false, "用例异常终止", "无异常", ex.GetType().Name + ": " + ex.Message, ex.StackTrace);
                }
            }

            try { app.Quit(); app.Dispose(); }
            catch { }

            Console.WriteLine();
            Console.WriteLine("══════════════════════════════════════════════");
            Console.WriteLine(" 执行用例: " + executed + " / " + cases.Count);
            if (env.Failures.Count == 0)
            {
                Console.WriteLine(" 全部通过 ✔（TC-01 ~ TC-18 全绿 = 黄金样本验收通过）");
            }
            else
            {
                Console.WriteLine(" 失败断言 " + env.Failures.Count + " 项 ✘：");
                foreach (string f in env.Failures) Console.WriteLine("   · " + f);
            }
            Console.WriteLine("══════════════════════════════════════════════");
            return env.Failures.Count;
        }
    }
}
