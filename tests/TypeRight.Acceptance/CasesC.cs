using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using NetOffice.WordApi;
using TypeRight.Formatting;
using TypeRight.Profiles;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>TC-15 ~ TC-18：方案闭环与性能。</summary>
    public static class CasesC
    {
        public static void TC15(TestEnv env)
        {
            env.Begin("TC-15 存为方案（4 条规则语义等价结构）");
            Document doc = env.OpenCopy("tc15");
            try
            {
                var journey = Journey.Run(env.App, doc, env, withFormulas: true);
                var profile = Journey.ToProfile(journey);

                var errors = ProfileV2Validator.Validate(profile);
                env.Check(errors.Count == 0, "方案符合 Schema v2", "0 错误", string.Join("；", errors), null);
                env.Check(profile.Rules.Count == 5, "规则数 = 5（模块/小节/正文/图注 + catchAll，语义等价结构）", "5", profile.Rules.Count.ToString(), null);
                env.Check(profile.ConvertFormulas, "公式开关 = true", "true", profile.ConvertFormulas.ToString(), null);

                var r1 = profile.Rules.FirstOrDefault(r => r.Order == 1);
                bool hasR1 = r1 != null; env.Check(hasR1, "规则1 存在", null, null, null); if (hasR1)
                {
                    env.Check(r1.Match.SizePt == 15, "规则1 match.sizePt = 15（模块标题）", "15", r1.Match.SizePt.ToString(), null);
                    env.Check(r1.Match.Bold == true, "规则1 match.bold = true", "true", (r1.Match.Bold ?? false).ToString(), null);
                    env.Check(r1.Match.Color == "#1A365D", "规则1 match.color = #1A365D", "#1A365D", r1.Match.Color, null);
                    env.Check(r1.Match.Ignore != null && r1.Match.Ignore.Contains("sizePt"), "规则1 ignore 含 sizePt", "含 sizePt", r1.Match.Ignore == null ? "(null)" : string.Join(",", r1.Match.Ignore), null);
                    env.Check(r1.Apply.EastAsiaFont == "楷体", "规则1 apply.eastAsiaFont = 楷体", "楷体", r1.Apply.EastAsiaFont, null);
                }

                var r2 = profile.Rules.FirstOrDefault(r => r.Order == 2);
                bool hasR2 = r2 != null; env.Check(hasR2, "规则2 存在", null, null, null); if (hasR2)
                {
                    env.Check(r2.Match.SizePt == 12.5, "规则2 match.sizePt = 12.5", "12.5", r2.Match.SizePt.ToString(), null);
                    env.Check(r2.Match.Color == "#2B6CB0", "规则2 match.color = #2B6CB0", "#2B6CB0", r2.Match.Color, null);
                    env.Check((r2.Match.Ignore == null || r2.Match.Ignore.Count == 0), "规则2 严格匹配（无 ignore）", "无", r2.Match.Ignore == null ? "无" : string.Join(",", r2.Match.Ignore), null);
                    env.Check(r2.Apply.EastAsiaFont == "黑体", "规则2 apply.eastAsiaFont = 黑体", "黑体", r2.Apply.EastAsiaFont, null);
                }

                // 规则 3/4 → 宋体，规则 5 = catchAll → 宋体（表格/★ 段收编，黄金取证）
                var songRules = profile.Rules.Where(r => r.Order == 3 || r.Order == 4).ToList();
                env.Check(songRules.Count == 2 && songRules.All(r => r.Apply.EastAsiaFont == "宋体"),
                    "规则3/4 apply.eastAsiaFont = 宋体", "2 条宋体规则", songRules.Count + " 条（" + string.Join(",", songRules.Select(r => r.Apply.EastAsiaFont)) + "）", null);
                var catchRule = profile.Rules.FirstOrDefault(r => r.Order == 5);
                env.Check(catchRule != null && catchRule.Match.CatchAll == true && catchRule.Apply.EastAsiaFont == "宋体",
                    "规则5 = catchAll → 宋体", "存在", catchRule == null ? "缺失" : (catchRule.Match.CatchAll == true ? "存在" : "非catchAll"), null);

                // 持久化
                var store = new ProfileStoreV2(env.OutputDir);
                store.Save(profile);
                env.Check(System.IO.File.Exists(System.IO.Path.Combine(env.OutputDir, profile.Id + ".json")),
                    "方案 JSON 落盘", "output/" + profile.Id + ".json", "缺失", null);
            }
            finally { doc.Close(false); }
        }

        public static void TC16(TestEnv env)
        {
            env.Begin("TC-16 应用方案（全新副本 → 结果与 TC-12 一致）");
            Document doc = env.OpenCopy("tc16");
            try
            {
                // 先在本副本上录制旅程并生成方案（避免依赖 TC-15 的执行顺序）
                var journey = Journey.Run(env.App, doc, env, withFormulas: true);
                var profile = Journey.ToProfile(journey);
                doc.Close(false);

                // 全新副本应用方案
                Document fresh = env.OpenCopy("tc16-apply");
                var report = ProfileApplier.Apply(env.App, fresh, profile);
                Console.WriteLine("  报告: " + string.Join(" / ", report.Lines));
                fresh.SaveAs2(env.CopyPath("tc16-applied"));
                fresh.Close(false);

                var hitLines = report.Lines.Where(l => l.StartsWith("规则")).ToList();
                int hit1 = ParseHit(hitLines.Count > 0 ? hitLines[0] : "");
                int hit2 = ParseHit(hitLines.Count > 1 ? hitLines[1] : "");
                int hit3 = ParseHit(hitLines.Count > 2 ? hitLines[2] : "");
                int hit4 = ParseHit(hitLines.Count > 3 ? hitLines[3] : "");
                int hit5 = ParseHit(hitLines.Count > 4 ? hitLines[4] : "");
                env.Check(hit1 == 8, "规则1（模块）命中 = 8", "8", hit1.ToString(), hitLines.Count > 0 ? hitLines[0] : "(无)");
                env.Check(hit2 == 16, "规则2（小节）命中 = 16", "16", hit2.ToString(), hitLines.Count > 1 ? hitLines[1] : "(无)");
                env.Check(hit3 == 64, "规则3（正文）命中 = 64", "64", hit3.ToString(), hitLines.Count > 2 ? hitLines[2] : "(无)");
                env.Check(hit4 == 7, "规则4（图注）命中 = 7", "7", hit4.ToString(), hitLines.Count > 3 ? hitLines[3] : "(无)");
                env.Check(hit5 == 50, "规则5（catchAll）命中 = 50（145-95，表格/★ 段）", "50", hit5.ToString(), hitLines.Count > 4 ? hitLines[4] : "(无)");
                env.Check(Math.Abs(report.FormulaConverted - 255) <= 2, "公式转换 = 255（±2）", "253~257", report.FormulaConverted.ToString(), null);
                env.Check(report.Unmatched == 0, "未命中 = 0", "0", report.Unmatched.ToString(), null);

                using (var xml = DocXml.Load(env.CopyPath("tc16-applied")))
                {
                    CasesB.CheckFontAnchors(env, xml, "TC-16");
                }
            }
            finally { try { doc.Close(false); } catch { } }
        }

        private static int ParseHit(string line)
        {
            int i = line.IndexOf("命中 ");
            if (i < 0) return -1;
            string rest = line.Substring(i + 3);
            string num = new string(rest.TakeWhile(char.IsDigit).ToArray());
            int v;
            return int.TryParse(num, out v) ? v : -1;
        }

        public static void TC17(TestEnv env)
        {
            env.Begin("TC-17 方案幂等（二次应用 0 改动）");
            Document doc = env.OpenCopy("tc17");
            try
            {
                var journey = Journey.Run(env.App, doc, env, withFormulas: true);
                var profile = Journey.ToProfile(journey);

                var report2 = ProfileApplier.Apply(env.App, doc, profile);
                env.Check(report2.TotalChanged == 0, "二次应用改动数 = 0", "0", report2.TotalChanged.ToString(), string.Join(" / ", report2.Lines));
                env.Check(report2.FormulaConverted == 0, "二次公式转换 = 0", "0", report2.FormulaConverted.ToString(), null);
            }
            finally { doc.Close(false); }
        }

        public static void TC18(TestEnv env)
        {
            env.Begin("TC-18 性能（匹配 <1s；黄金旅程 <10s）");
            Document doc = env.OpenCopy("tc18");
            try
            {
                // ① 同类选择匹配（95 段 + 11 表）
                Paragraph module = env.FindParagraph(doc, "模块一");
                var sig = FormatSignatureReader.Read(module);
                var sw = Stopwatch.StartNew();
                var hit = SimilarSelector.Match(doc, sig, FormatSignature.IgnoreDims.None);
                sw.Stop();
                int totalParas = doc.Paragraphs.Count;
                Console.WriteLine(string.Format("  匹配: {0} 段命中 / {1}ms（文档段落总数 {2}，进程外 DCOM 环境）",
                    hit.Count, sw.ElapsedMilliseconds, totalParas));
                env.Check(sw.ElapsedMilliseconds < 1000, "同类选择匹配 < 1s（进程外 DCOM 参考口径）",
                    "<1000ms", sw.ElapsedMilliseconds + "ms",
                    "注：验收以进程内为准（addin-log 实测 1000 段/532ms）");

                // ② 黄金旅程全程
                Document journeyDoc = env.OpenCopy("tc18-journey");
                try
                {
                    var journey = Journey.Run(env.App, journeyDoc, env, withFormulas: true);
                    Console.WriteLine(string.Format("  旅程: 匹配 {0}/{1}/{2}/{3} → 公式 {4}，总耗时 {5}ms（进程外 DCOM）",
                        journey.Round1Hits, journey.Round2Hits, journey.Round3Hits, journey.Round4Hits,
                        journey.Formula != null ? journey.Formula.Converted.ToString() : "-", journey.TotalMs));
                    env.Check(journey.TotalMs < 10000, "黄金旅程全程 < 10s", "<10000ms", journey.TotalMs + "ms", "进程外 DCOM 环境");
                }
                finally { journeyDoc.Close(false); }
            }
            finally { doc.Close(false); }
        }
    }
}
