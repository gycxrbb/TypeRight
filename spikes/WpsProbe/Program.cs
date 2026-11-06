// ============================================================================
// WPS COM 兼容性探针（独立工程，与 TypeRight 产品代码零耦合）
// 目的：回答"WPS 能否承载 TypeRight 的三大功能"：
//   A 匹配+批量格式（同类选择）——依赖 Paragraphs/Range/Font 读写/表格
//   B 公式转换——依赖 OMaths / FormattedText / InsertFile / 隐藏文档
//   C 侧车 UI——依赖 Window.Hwnd / Selection / SaveAs
// 方法：对 KWPS.Application 纯 late-binding(dynamic) 逐 API 体检，PASS/FAIL/WARN 落报告。
// ============================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace WpsProbe
{
    internal static class Program
    {
        private static readonly List<string[]> Results = new List<string[]>(); // group, mark, name, detail
        private static int _pass, _fail, _warn;

        private static void Pass(string g, string n, string d = "") { _pass++; Results.Add(new[] { g, "PASS", n, d }); }
        private static void Fail(string g, string n, string d) { _fail++; Results.Add(new[] { g, "FAIL", n, d }); }
        private static void Warn(string g, string n, string d) { _warn++; Results.Add(new[] { g, "WARN", n, d }); }

        private static string Ex(Exception ex)
        {
            var e = ex.InnerException ?? ex;
            return e.GetType().Name + ": " + e.Message;
        }

        private static string Describe(object v)
        {
            if (v == null) return "null";
            string s = v.ToString();
            if (s.Length > 64) s = s.Substring(0, 64) + "…";
            return s.Replace("\r", "\\r").Replace("\n", "\\n");
        }

        /// <summary>单项探测：返回值非 null=通过；异常=FAIL（或 WARN）。</summary>
        private static object Get(string g, string name, Func<object> f, bool required = true)
        {
            try { var v = f(); Pass(g, name, Describe(v)); return v; }
            catch (Exception ex) { if (required) Fail(g, name, Ex(ex)); else Warn(g, name, Ex(ex)); return null; }
        }

        private static void Do(string g, string name, Action a, bool required = true)
        {
            try { a(); Pass(g, name); }
            catch (Exception ex) { if (required) Fail(g, name, Ex(ex)); else Warn(g, name, Ex(ex)); }
        }

        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("════ WPS COM 兼容性探针 ════");

            // ── 定位仓库根（spikes/WpsProbe/bin/Debug → 上溯到有 src/ 的目录）──
            string repo = AppDomain.CurrentDomain.BaseDirectory;
            while (repo != null && !Directory.Exists(Path.Combine(repo, "src"))) repo = Path.GetDirectoryName(repo);
            if (repo == null) { Console.WriteLine("[FATAL] 找不到仓库根（src/ 不存在）"); return 2; }
            string outDir = Path.Combine(repo, "spikes", "WpsProbe", "out");
            Directory.CreateDirectory(outDir);

            string lecture = Path.Combine(repo, "tests", "初三数学代数终极讲义_一元二次方程与二次函数大一统.docx");
            if (args.Length > 0) lecture = args[0];
            // 含 255 处公式的转换产物（OMML 往返测试源）
            string formulaDoc = new DirectoryInfo(Path.Combine(repo, "tests", "output"))
                .EnumerateFiles("tc10-out-*.docx").OrderByDescending(f => f.LastWriteTime)
                .Select(f => f.FullName).FirstOrDefault();
            Console.WriteLine("资产: " + (File.Exists(lecture) ? "OK" : "缺失!") + " " + Path.GetFileName(lecture));
            Console.WriteLine("公式源: " + (formulaDoc != null && File.Exists(formulaDoc) ? "OK" : "缺失!") + " " + Path.GetFileName(formulaDoc ?? "-"));

            // ── 环境情报（不动 WPS）──
            ProbeWordApplicationOwner();
            bool wpsPreexisting = ProbeWpsProcesses("启动前");

            // ── 激活 WPS ──
            const string G = "0-激活";
            dynamic app = null;
            try
            {
                Type t = null; string progId = null;
                foreach (string pid in new[] { "KWPS.Application", "wps.Application", "WPS.Application" })
                {
                    t = Type.GetTypeFromProgID(pid);
                    if (t != null) { progId = pid; break; }
                }
                if (t == null) { Fail(G, "ProgID 解析", "KWPS/wps/WPS.Application 均未注册"); return 1; }
                app = Activator.CreateInstance(t);
                Pass(G, "CreateInstance(" + progId + ")");
            }
            catch (Exception ex) { Fail(G, "CreateInstance", Ex(ex)); return 1; }

            Do(G, "Visible=false", () => { app.Visible = false; });
            Do(G, "DisplayAlerts=0", () => { app.DisplayAlerts = 0; }, required: false);
            Get(G, "app.Version", () => app.Version);
            Get(G, "app.Build", () => app.Build, required: false);
            Get(G, "app.Path", () => app.Path, required: false);

            dynamic lectureDoc = null, scratch = null, hidden = null;
            try
            {
                // ════════ A：文档与段落（匹配的地基）════════
                const string GA = "A-文档/段落";
                string lectureCopy = Path.Combine(outDir, "lecture-copy.docx");
                File.Copy(lecture, lectureCopy, true);
                File.SetAttributes(lectureCopy, FileAttributes.Normal);
                lectureDoc = app.Documents.Open(lectureCopy, Type.Missing, true);
                Pass(GA, "Documents.Open(ReadOnly)", lectureDoc.FullName);

                int paraCount = (int)(Get(GA, "Paragraphs.Count", () => lectureDoc.Paragraphs.Count) ?? 0);

                // 枚举器（TypeRight 全程用枚举器，绝不用索引器——O(n) 大坑）
                int enumCount = 0;
                try
                {
                    foreach (dynamic p in lectureDoc.Paragraphs) { enumCount++; if (enumCount >= 2000) break; }
                    Pass(GA, "Paragraphs foreach 枚举", enumCount + " 段");
                }
                catch (Exception ex) { Fail(GA, "Paragraphs foreach 枚举", Ex(ex)); }

                dynamic p1 = lectureDoc.Paragraphs[1];
                object rStart = Get(GA, "Paragraph.Range.Start", () => p1.Range.Start);
                object rEnd = Get(GA, "Paragraph.Range.End", () => p1.Range.End);
                string p1Text = (string)Get(GA, "Paragraph.Range.Text", () => p1.Range.Text, required: false);
                Get(GA, "doc.Content.End", () => lectureDoc.Content.End);
                Get(GA, "doc.Range(a,b).Text", () => lectureDoc.Range(0, Math.Min(20, (int)lectureDoc.Content.End)).Text, required: false);

                int tableCount = (int)(Get(GA, "Tables.Count", () => lectureDoc.Tables.Count) ?? -1);
                if (tableCount > 0)
                {
                    Get(GA, "Tables[1].Range.Start/End", () =>
                    {
                        var t1 = lectureDoc.Tables[1];
                        return "start=" + t1.Range.Start + " end=" + t1.Range.End;
                    });
                    Get(GA, "Tables[1].Cell(1,1).Range.Text", () => lectureDoc.Tables[1].Cell(1, 1).Range.Text, required: false);
                }
                // 我们在产品里不信任 Information(wdWithInTable)（进程内假值前科），此处纯情报
                Get(GA, "Range.Information(21) 情报", () => lectureDoc.Range(0, 1).Information(21), required: false);

                // ════════ B：字体读写（签名的原料）════════
                const string GB = "B-字体读";
                if (paraCount > 0 && p1Text != null && p1Text.Length > 1)
                {
                    int s0 = (int)rStart;
                    dynamic fc = lectureDoc.Range(s0, s0 + 1).Font;
                    Get(GB, "Font.NameFarEast(单字符)", () => fc.NameFarEast);
                    Get(GB, "Font.NameAscii(单字符)", () => fc.NameAscii, required: false);
                    Get(GB, "Font.NameOther(单字符)", () => fc.NameOther, required: false);
                    Get(GB, "Font.Name(单字符)", () => fc.Name);
                    Get(GB, "Font.Size", () => fc.Size);
                    Get(GB, "Font.Bold", () => fc.Bold);
                    Get(GB, "Font.Italic", () => fc.Italic, required: false);
                    Get(GB, "Font.Underline", () => fc.Underline, required: false);
                    Get(GB, "Font.Color", () => fc.Color);
                    Get(GB, "Font.HighlightColorIndex", () => fc.HighlightColorIndex, required: false);
                }
                // 段级整读（中段锚点的读法）
                Get(GB, "段 Range.Font.NameFarEast", () => p1.Range.Font.NameFarEast, required: false);
                Get(GB, "段 Range.ParagraphFormat.Alignment", () => p1.Range.ParagraphFormat.Alignment, required: false);

                // ════════ 性能：签名匹配热路径模拟（段落文本 + 5 锚点×4 属性 ≈ 20 COM 读/段）════════
                if (enumCount > 0)
                {
                    var sw = Stopwatch.StartNew();
                    int reads = 0, paras = 0; double chars = 0;
                    foreach (dynamic p in lectureDoc.Paragraphs)
                    {
                        dynamic r = p.Range;
                        string txt = (string)r.Text; if (txt == null) continue;
                        paras++; chars += txt.Length;
                        int len = txt.Length;
                        if (len > 1)
                        {
                            int step = Math.Max(1, (len - 1) / 5);
                            for (int k = 0; k < 5; k++)
                            {
                                int pos = (int)r.Start + Math.Min(len - 1, k * step);
                                dynamic fnt = lectureDoc.Range(pos, pos + 1).Font;
                                reads++;
                                var a = fnt.NameFarEast; var b = fnt.Size; var c = fnt.Bold; var d = fnt.Color;
                                reads += 3;
                            }
                        }
                        if (paras >= 500) break;
                    }
                    sw.Stop();
                    double perCall = reads + paras > 0 ? sw.Elapsed.TotalMilliseconds / (reads + paras) : -1;
                    Warn("B-字体读", "热路径计时(情报)", string.Format(
                        "{0} 段/{1} 次 COM 读 {2:N0} ms（{3:N3} ms/次；Word 进程内基线 ~0.5s/千段）",
                        paras, reads + paras, sw.Elapsed.TotalMilliseconds, perCall));
                }

                // ════════ C：写入路径（批量格式 + 高亮 + 撤销）════════
                const string GC = "C-格式写";
                scratch = app.Documents.Add(Type.Missing, Type.Missing, Type.Missing, true);
                Pass(GC, "Documents.Add");
                Do(GC, "Range.Text 赋值", () =>
                {
                    dynamic r = scratch.Range(0, 0);
                    r.Text = "第一段普通文本段落，用于格式写入探针。\r第二段公式样例 x²+1=0。\r";
                });
                dynamic para1 = scratch.Paragraphs[1];
                dynamic tgt = scratch.Range(para1.Range.Start, para1.Range.End - 1); // 排除段落标记（高亮扩散坑）

                Do(GC, "ScreenUpdating=false/true", () => { app.ScreenUpdating = false; app.ScreenUpdating = true; });

                Do(GC, "Font.NameFarEast=楷体", () => { tgt.Font.NameFarEast = "楷体"; });
                Do(GC, "Font.Name=Times New Roman", () => { tgt.Font.Name = "Times New Roman"; });
                Do(GC, "Font.Size=15", () => { tgt.Font.Size = 15.0; });
                Do(GC, "Font.Bold=-1", () => { tgt.Font.Bold = -1; });
                Do(GC, "Font.Italic=-1", () => { tgt.Font.Italic = -1; }, required: false);
                Do(GC, "Font.Underline=1", () => { tgt.Font.Underline = 1; }, required: false);
                Do(GC, "Font.Color=0x1A365D", () => { tgt.Font.Color = 0x5D361A; }); // BGR
                Do(GC, "Range.HighlightColorIndex=7(黄)", () => { tgt.HighlightColorIndex = 7; });
                Do(GC, "Range.HighlightColorIndex=0(清)", () => { tgt.HighlightColorIndex = 0; });
                Do(GC, "ParagraphFormat.Alignment=1(中)", () => { tgt.ParagraphFormat.Alignment = 1; }, required: false);

                // 回读核验（用新鲜区间，避免任何缓存假象）
                Get(GC, "回读 NameFarEast=楷体?", () => scratch.Range(para1.Range.Start, para1.Range.Start + 2).Font.NameFarEast, required: false);
                Get(GC, "回读 Size=15?", () => scratch.Range(para1.Range.Start, para1.Range.Start + 2).Font.Size, required: false);
                Get(GC, "回读 Bold=-1?", () => scratch.Range(para1.Range.Start, para1.Range.Start + 2).Font.Bold, required: false);

                // 撤销：UndoRecord（Word 2013+）在 WPS 上预期不可用 → 情报项；doc.Undo 兜底可用性
                Get(GC, "app.UndoRecord（情报）", () => app.UndoRecord, required: false);
                Do(GC, "doc.Undo() 兜底", () => { scratch.Undo(); }, required: false);
                Get(GC, "Undo 后 Size 回退?", () => scratch.Range(para1.Range.Start, para1.Range.Start + 2).Font.Size, required: false);

                // ════════ D：公式管线原语 ════════
                const string GD = "D-公式";
                object om = Get(GD, "doc.OMaths（对象模型公式）", () => lectureDoc.OMaths, required: false);
                bool hasOmaths = om != null;
                if (hasOmaths)
                {
                    Get(GD, "OMaths.Count", () => lectureDoc.OMaths.Count, required: false);
                    Do(GD, "OMaths.Add(range)+BuildUp", () =>
                    {
                        dynamic p2 = scratch.Paragraphs[2];
                        dynamic xr = scratch.Range(p2.Range.Start + 5, p2.Range.Start + 10);
                        dynamic m = scratch.OMaths.Add(xr);
                        m.BuildUp();
                    }, required: false);
                }
                else
                {
                    Fail(GD, "doc.OMaths", "WPS 无 OMaths 对象模型 → 走 OMML 文档注入路线（下续探测）");
                }

                // 隐藏文档 + InsertFile（三阶段管线的装载原语）
                hidden = app.Documents.Add(Type.Missing, Type.Missing, Type.Missing, false);
                Pass(GD, "Documents.Add(Visible:=false) 隐藏草稿");
                bool inserted = false;
                if (formulaDoc != null)
                {
                    Do(GD, "InsertFile(含公式 docx)", () => { hidden.Range(0, 0).InsertFile(formulaDoc); });
                    inserted = true;
                    Get(GD, "插入后 Content.End", () => hidden.Content.End, required: false);
                }
                else Fail(GD, "InsertFile", "无公式测试资产");

                // FormattedText 复制（批量注入的性能主路径）
                if (inserted)
                {
                    Do(GD, "target.FormattedText = src.FormattedText", () =>
                    {
                        dynamic src = hidden.Range(0, Math.Min(120, (int)hidden.Content.End - 1));
                        dynamic dst = scratch.Range(scratch.Content.End - 1, scratch.Content.End - 1);
                        dst.FormattedText = src.FormattedText;
                    });
                    Get(GD, "复制后段落计数", () => scratch.Paragraphs.Count, required: false);
                }

                // OMML 往返：隐藏文档另存 → 检查 m:oMath 是否存活
                if (inserted)
                {
                    string saved = Path.Combine(outDir, "wps-roundtrip.docx");
                    try
                    {
                        try { hidden.SaveAs2(saved, 12); }   // 12 = wdFormatXMLDocument
                        catch { hidden.SaveAs(saved, 12); }
                        int hits = CountOmath(saved);
                        if (hits > 0) Pass(GD, "OMML 往返存活", saved + " 内 m:oMath × " + hits);
                        else Fail(GD, "OMML 往返存活", "另存 docx 中未检出 m:oMath");
                    }
                    catch (Exception ex) { Fail(GD, "OMML 往返存活", Ex(ex)); }

                    if (hasOmaths) Get(GD, "往返后 OMaths.Count", () => hidden.OMaths.Count, required: false);
                }
                Get(GD, "SaveAs2 情报", () => { hidden.SaveAs2(Path.Combine(outDir, "saveas2-probe.docx"), 12); return "OK"; }, required: false);

                // ════════ E：UI/侧车架构原语 ════════
                const string GE = "E-UI/侧车";
                Get(GE, "ActiveWindow.Hwnd", () => app.ActiveWindow.Hwnd, required: false);
                Get(GE, "ActiveWindow.Caption", () => app.ActiveWindow.Caption, required: false);
                Do(GE, "Selection.SetRange", () => { app.Selection.SetRange(0, 5); }, required: false);
                Get(GE, "Selection.Text", () => app.Selection.Text, required: false);

                object addinCount = Get(GE, "COMAddIns.Count（WPS 是否有 COM 加载项体系）", () => app.COMAddIns.Count, required: false);
                if (addinCount != null)
                {
                    try
                    {
                        int n = Math.Min(5, (int)addinCount);
                        var names = new StringBuilder();
                        for (int i = 1; i <= n; i++) names.Append(app.COMAddIns.Item(i).ProgId).Append(" ");
                        Pass(GE, "COMAddIns 枚举", names.ToString());
                    }
                    catch (Exception ex) { Warn(GE, "COMAddIns 枚举", Ex(ex)); }
                }
            }
            catch (Exception ex)
            {
                Fail("Z-整体", "未预期异常", Ex(ex));
            }
            finally
            {
                // 收尾：只关自己开的文档；若启动前就有 wps.exe（用户的会话）则不动 Quit
                foreach (dynamic d in new object[] { lectureDoc, scratch, hidden })
                {
                    try { if (d != null) d.Close(false); } catch { }
                }
                if (!wpsPreexisting)
                {
                    try { app.Quit(); } catch { }
                }
                else Warn("0-激活", "跳过 Quit（检测到预存 wps.exe，请自行关闭探针拉起的实例）", "");
            }

            // wps.exe 僵尸检查
            ProbeWpsProcesses("收尾后");

            // ── 报告 ──
            var sb = new StringBuilder();
            sb.AppendLine("WPS COM 兼容性探针报告  " + DateTime.Now.ToString("s"));
            sb.AppendLine(new string('─', 72));
            string lastGroup = null;
            foreach (var r in Results)
            {
                if (r[0] != lastGroup) { sb.AppendLine(); sb.AppendLine("【" + r[0] + "】"); lastGroup = r[0]; }
                sb.AppendFormat("  [{0}] {1}{2}" + Environment.NewLine,
                    r[1], r[2], string.IsNullOrEmpty(r[3]) ? "" : " —— " + r[3]);
            }
            sb.AppendLine();
            sb.AppendLine(new string('─', 72));
            sb.AppendFormat("总计: PASS {0} / FAIL {1} / WARN {2}", _pass, _fail, _warn).AppendLine();
            string reportPath = Path.Combine(repo, "spikes", "WpsProbe", "report.txt");
            File.WriteAllText(reportPath, sb.ToString(), new UTF8Encoding(true));
            Console.WriteLine();
            Console.WriteLine(sb.ToString());
            Console.WriteLine("报告已写入: " + reportPath);
            return _fail == 0 ? 0 : 1;
        }

        private static int CountOmath(string docxPath)
        {
            using (var z = ZipFile.OpenRead(docxPath))
            {
                var entry = z.GetEntry("word/document.xml");
                if (entry == null) return -1;
                using (var sr = new StreamReader(entry.Open(), Encoding.UTF8))
                {
                    string xml = sr.ReadToEnd();
                    int n = 0, i = 0;
                    while ((i = xml.IndexOf("<m:oMath", i, StringComparison.Ordinal)) >= 0) { n++; i += 8; }
                    return n;
                }
            }
        }

        private static void ProbeWordApplicationOwner()
        {
            try
            {
                var cls = Registry.ClassesRoot.OpenSubKey(@"Word.Application\CLSID")?.GetValue(null) as string;
                string srv = null;
                if (cls != null)
                    srv = Registry.ClassesRoot.OpenSubKey(@"CLSID\" + cls + @"\LocalServer32")?.GetValue(null) as string;
                Warn("0-激活", "Word.Application 归属（情报）", "CLSID=" + cls + " Server=" + srv);
            }
            catch (Exception ex) { Warn("0-激活", "Word.Application 归属", Ex(ex)); }
        }

        private static bool ProbeWpsProcesses(string tag)
        {
            try
            {
                var ps = Process.GetProcessesByName("wps");
                if (ps.Length > 0)
                    Warn("0-激活", "wps.exe 进程（" + tag + "）", ps.Length + " 个存活（PID: " + string.Join(",", ps.Take(4).Select(p => p.Id)) + "）");
                else
                    Pass("0-激活", "wps.exe 进程（" + tag + "）", "无");
                return ps.Length > 0;
            }
            catch { return false; }
        }
    }
}
