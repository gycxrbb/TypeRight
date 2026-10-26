using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Classification;
using TypeRight.Formatting;
using TypeRight.Profiles;
using TypeRight.Roles;
using TypeRight.Scanning;

namespace TypeRight.SmokeTest
{
    /// <summary>
    /// 端到端冒烟测试（真实 Word 进程）：
    /// 构造脏格式文档 → 扫描分类（断言 R1-R10）→ 应用排版（断言样式/字体/强调保留/公式）
    /// → 单次 Undo 完全还原 → 二次应用 0 改动（幂等）。
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
                Console.WriteLine("TypeRight 冒烟测试（真实 Word）");
                return Run(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[FATAL] " + ex);
                return 2;
            }
        }

        private static int Run(string[] args)
        {
            bool keepWordVisible = args.Contains("--keep");

            Application app = new Application();
            app.Visible = false;
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            Document doc = null;
            try
            {
                doc = app.Documents.Add();
                BuildMessyDocument(doc);

                // ---- 1. 扫描 + 分类 ----
                Header("1. 扫描与角色识别");
                var scanner = new DocumentScanner(app);
                var scan = scanner.Scan(doc);
                RoleClassifier.Classify(scan.Features);
                Console.WriteLine("  扫描段落数: " + scan.Features.Count + "，用时 " + scan.Duration.TotalMilliseconds.ToString("0") + "ms");
                foreach (var f in scan.Features)
                    Console.WriteLine(string.Format("    #{0,-2} R{1,-2} {2,-8} {3}", f.Index, f.RulePriority, f.Role, f.DisplayText(30)));

                var fx = scan.Features;
                CheckRole(fx, "一元二次方程讲义", ParagraphRole.Title);
                CheckRole(fx, "第一章", ParagraphRole.Heading1);
                CheckRole(fx, "1.1 ", ParagraphRole.Heading2);
                CheckRole(fx, "1.1.2", ParagraphRole.Heading3);
                CheckRole(fx, "图 1", ParagraphRole.Caption);
                CheckRoleStartsWith(fx, "x₁²", ParagraphRole.FormulaLine);
                var tableCellPara = fx.FirstOrDefault(p => p.InTable);
                Check(tableCellPara != null, "扫描包含表格内段落");
                var bodyPara = fx.FirstOrDefault(p => p.Role == ParagraphRole.Body && p.Text.StartsWith("对于方程"));
                Check(bodyPara != null, "正文段落识别（含内联公式文本）");
                Check(fx.Any(p => p.Role == ParagraphRole.Skip), "空段被跳过（R1）");

                // ---- 2. 应用排版 ----
                Header("2. 应用排版");
                var profile = ProfileStore.BuiltInProfile;
                var assignments = new Dictionary<int, ParagraphRole>();
                foreach (var f in scan.Features) assignments[f.Index] = f.Role;

                // 撤销基线：应用前的 Normal 样式定义
                Style normalBefore = doc.Styles[(int)WdBuiltinStyle.wdStyleNormal];
                string preNormalFont = normalBefore.Font.NameFarEast;
                float preNormalSize = normalBefore.Font.Size;

                var request = new FormatRequest
                {
                    Profile = profile,
                    RoleAssignments = assignments,
                    MountStyles = true,
                    ConvertFormulas = true,
                    Selection = null
                };
                var stats = new FormatApplier().Apply(app, doc, request);
                Console.WriteLine("  " + stats.Summary());
                Check(stats.ChangedParagraphs > 0, "发生实际改动（ChangedParagraphs=" + stats.ChangedParagraphs + "）");

                // 样式定义
                Style normal = doc.Styles[(int)WdBuiltinStyle.wdStyleNormal];
                Style heading1 = doc.Styles[(int)WdBuiltinStyle.wdStyleHeading1];
                Check(Math.Abs(normal.Font.Size - 10.5f) < 0.05, "Normal 样式字号 = 10.5（实际 " + normal.Font.Size + "）");
                Check(string.Equals(normal.Font.NameFarEast, "宋体", StringComparison.Ordinal), "Normal 中文名 = 宋体（实际 " + normal.Font.NameFarEast + "）");
                Check(string.Equals(normal.Font.NameAscii, "Times New Roman", StringComparison.OrdinalIgnoreCase), "Normal 西文名 = Times New Roman");
                Check(string.Equals(heading1.Font.NameFarEast, "黑体", StringComparison.Ordinal), "标题1 中文名 = 黑体（实际 " + heading1.Font.NameFarEast + "）");
                Check(heading1.Font.Bold == -1, "标题1 加粗");

                // 段落挂载
                var titlePara = FindParagraph(doc, "一元二次方程讲义");
                Check(titlePara != null && titlePara.Style is Style && StyleIs(titlePara, WdBuiltinStyle.wdStyleTitle), "主标题挂载 Title 样式");
                var h1Para = FindParagraph(doc, "第一章");
                Check(h1Para != null && StyleIs(h1Para, WdBuiltinStyle.wdStyleHeading1), "第一章挂载 Heading 1 样式");
                Check(h1Para != null && Convert.ToInt32(h1Para.OutlineLevel) == (int)WdOutlineLevel.wdOutlineLevel1, "一级标题大纲级别 = 1（导航窗格可见）");

                // 直接字体污染被清除/覆盖（标题段落视觉正确）
                if (titlePara != null)
                    Check(Math.Abs(titlePara.Range.Font.Size - 16f) < 0.05f, "主标题有效字号 = 16（实际 " + titlePara.Range.Font.Size + "）");

                // 保留强调：正文段中手动加粗的“判别式”仍是加粗
                var bodyWord = FindRange(doc, "判别式与求根公式的应用");
                Check(bodyWord != null && bodyWord.Font.Bold == -1, "局部加粗强调被保留");

                // 公式转换
                int omathCount = doc.OMaths.Count;
                Console.WriteLine("  公式对象数: " + omathCount + "（转换 " + stats.FormulasConverted + "，失败 " + stats.FormulasFailed + "）");
                Check(omathCount >= 5, "公式对象 ≥ 5");
                Check(stats.FormulasFailed == 0, "公式转换零失败");
                // 根式公式文本（BuildUp 后线性文本仍在数学区）
                // BuildUp 后字母转为数学斜体字符（𝑎 = U+1D44E），验证公式内容仍在文本流中
                string docText = doc.Content.Text ?? "";
                Check(docText.Contains("\U0001D44E") || docText.Contains("2a"), "根式公式内容保留（数学斜体 𝑎 存在）");

                // 混排规则（内置方案默认关闭，跳过）

                // ---- 3. 单次撤销 ----
                Header("3. 单次 Ctrl+Z 完全还原");
                bool undone = doc.Undo();
                Check(undone, "Document.Undo() 成功");
                Check(doc.OMaths.Count == 0, "撤销后公式对象清零（实际 " + doc.OMaths.Count + "）");
                var titleAfter = FindParagraph(doc, "一元二次方程讲义");
                Check(titleAfter != null && Math.Abs(titleAfter.Range.Font.Size - 20f) < 0.05f,
                    "撤销后主标题恢复 20pt（实际 " + (titleAfter == null ? "?" : titleAfter.Range.Font.Size.ToString()) + "）");
                Style normalAfter = doc.Styles[(int)WdBuiltinStyle.wdStyleNormal];
                Check(string.Equals(normalAfter.Font.NameFarEast, preNormalFont, StringComparison.Ordinal) &&
                      Math.Abs(normalAfter.Font.Size - preNormalSize) < 0.05f,
                    "撤销后 Normal 样式定义回退（字体 " + normalAfter.Font.NameFarEast + " / 字号 " + normalAfter.Font.Size +
                    "，基线 " + preNormalFont + " / " + preNormalSize + "）");

                // ---- 4. 幂等 ----
                Header("4. 幂等：二次应用 0 改动");
                // 撤销一次（恢复原始文档）→ 重新应用 → 再次应用验证 0 改动
                FormatApplier.DebugLog = false;
                var stats2 = new FormatApplier().Apply(app, doc, request);
                var stats3 = new FormatApplier().Apply(app, doc, request);
                Console.WriteLine("  第二次: " + stats3.Summary());
                Check(stats3.ChangedParagraphs == 0, "二次应用 ChangedParagraphs = 0（实际 " + stats3.ChangedParagraphs + "）");
                Check(stats3.FormulasConverted == 0, "二次应用公式转换 = 0（实际 " + stats3.FormulasConverted + "）");
                Check(stats3.StyleDefinitionsUpdated == 0, "二次应用样式定义写入 = 0（实际 " + stats3.StyleDefinitionsUpdated + "）");

                // ---- 5. 导出结果供人工检查 ----
                string outDir = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "out");
                System.IO.Directory.CreateDirectory(outDir);
                string outPath = System.IO.Path.Combine(outDir, "smoke-result.docx");
                try
                {
                    doc.SaveAs2(outPath);
                    Console.WriteLine();
                    Console.WriteLine("结果文档: " + outPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("保存结果文档失败: " + ex.Message);
                }
            }
            finally
            {
                try { if (doc != null) { doc.Close(false); doc.Dispose(); } }
                catch { }
                try { if (!keepWordVisible) { app.Quit(); app.Dispose(); } }
                catch { }
            }

            Console.WriteLine();
            Console.WriteLine(_failures == 0 ? "全部通过 ✔" : "失败 " + _failures + " 项 ✘");
            return _failures == 0 ? 0 : 1;
        }

        // ---- 断言助手 ----

        private static void CheckRole(List<ParagraphFeatures> fx, string startsWith, ParagraphRole role)
        {
            var f = fx.FirstOrDefault(p => p.Text.StartsWith(startsWith, StringComparison.Ordinal));
            Check(f != null && f.Role == role, "「" + startsWith + "…」→ " + role.ToLabel() +
                (f == null ? "（未找到段落）" : "（实际 " + f.Role.ToLabel() + "，规则 R" + f.RulePriority + "）"));
        }

        private static void CheckRoleStartsWith(List<ParagraphFeatures> fx, string startsWith, ParagraphRole role)
        {
            var f = fx.FirstOrDefault(p => p.Text.StartsWith(startsWith, StringComparison.Ordinal));
            Check(f != null && f.Role == role, "「" + startsWith + "…」→ " + role.ToLabel() +
                (f == null ? "（未找到段落）" : "（实际 " + f.Role.ToLabel() + "）"));
        }

        // ---- 文档构造 ----

        private static void BuildMessyDocument(Document doc)
        {
            Range rng = doc.Range(0, 0);

            // 主标题：20pt 加粗 直接格式
            AddPara(rng, "一元二次方程讲义", r => { r.Font.Size = 20; r.Font.Bold = -1; });
            // 空段
            AddPara(rng, "", null);
            // 一级标题（纯文本，无格式）
            AddPara(rng, "第一章 方程与函数", null);
            // 二级标题
            AddPara(rng, "1.1 一元二次方程的定义", null);
            // 正文（含内联公式文本 + 局部加粗强调）
            AddPara(rng, "对于方程 ax² + bx + c = 0，其中 a≠0，我们称 x₁ 和 x₂ 为它的两个根。", null);
            AddPara(rng, "只含一个未知数，且未知数的最高次数是 2 的整式方程叫作一元二次方程。", null);
            // 三级标题
            AddPara(rng, "1.1.2 根的判别", null);
            // 正文：内嵌根式公式
            AddPara(rng, "将各项系数代入求根公式(-b±√(b²-4ac))/(2a)即可求出两根。", null);
            // 独立公式行
            AddPara(rng, "x₁² + x₂² = (x₁ + x₂)² - 2x₁x₂", null);
            // 图注
            AddPara(rng, "图 1 一元二次函数的图像", null);
            // 模块式标题
            AddPara(rng, "模块二：例题精讲", null);
            // 正文，其中有手动加粗的强调词
            AddPara(rng, "例 1：解方程 x² - 5x + 6 = 0。解题关键是判别式与求根公式的应用。", r =>
            {
                var word = r.Duplicate;
                int idx = r.Text.IndexOf("判别式与求根公式的应用", StringComparison.Ordinal);
                if (idx >= 0)
                {
                    word.Start = r.Start + idx;
                    word.End = word.Start + "判别式与求根公式的应用".Length;
                    word.Font.Bold = -1;
                }
            });
            // 表格（首行加粗 = 表头）
            AddTable(doc, rng);
        }

        private static void AddPara(Range rng, string text, Action<Range> format)
        {
            rng.Collapse(WdCollapseDirection.wdCollapseEnd);
            rng.InsertAfter(text);
            rng.InsertParagraphAfter();
            // 定位刚插入的段落
            var paraRange = rng.Duplicate;
            paraRange.Start = paraRange.End - (text.Length + 1);
            paraRange.End = paraRange.Start + text.Length;
            if (format != null) format(paraRange);
        }

        private static void AddTable(Document doc, Range rng)
        {
            rng.Collapse(WdCollapseDirection.wdCollapseEnd);
            rng.InsertParagraphAfter();
            var anchor = doc.Range(rng.End - 1, rng.End - 1);
            Table table = doc.Tables.Add(anchor, 3, 2);
            table.Cell(1, 1).Range.Text = "参数";
            table.Cell(1, 2).Range.Text = "含义";
            var headerRow = table.Rows[1].Range;
            headerRow.Font.Bold = -1;
            table.Cell(2, 1).Range.Text = "a";
            table.Cell(2, 2).Range.Text = "二次项系数，a≠0";
            table.Cell(3, 1).Range.Text = "Δ";
            table.Cell(3, 2).Range.Text = "b² - 4ac，决定根的个数";
        }

        // ---- Word 查找助手 ----

        private static Paragraph FindParagraph(Document doc, string startsWith)
        {
            Paragraphs ps = doc.Paragraphs;
            for (int i = 1; i <= ps.Count; i++)
            {
                try
                {
                    string t = ps[i].Range.Text ?? "";
                    if (t.TrimEnd('\r', '\a').StartsWith(startsWith, StringComparison.Ordinal)) return ps[i];
                }
                catch { }
            }
            return null;
        }

        private static bool StyleIs(Paragraph p, WdBuiltinStyle builtin)
        {
            try
            {
                Document doc = (Document)p.Range.Document;
                Style target = doc.Styles[(int)builtin];
                object cur = p.Style;
                if (cur is Style) return string.Equals(((Style)cur).NameLocal, target.NameLocal, StringComparison.Ordinal);
                return string.Equals(Convert.ToString(cur), target.NameLocal, StringComparison.Ordinal);
            }
            catch { return false; }
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

        private static bool ContainsText(Document doc, string text)
        {
            return FindRange(doc, text) != null;
        }
    }
}
