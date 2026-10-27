using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Formatting
{
    /// <summary>
    /// 示例/测试文档生成器（PRD v2 黄金样本）：用直接格式构造出明确的签名分组——
    /// 主标题(22pt 粗 深蓝)、模块标题×2(15pt 粗 深蓝)、小节标题×3(12.5pt 粗 中蓝)、
    /// 正文若干(10.5pt)、图注(9pt 灰)、表格。同类选择的命中数可精确断言。
    /// </summary>
    public static class SampleDocBuilder
    {
        public static Document Build(Application app)
        {
            if (app == null) throw new ArgumentNullException("app");
            Document doc = app.Documents.Add();

            Range rng = doc.Range(0, 0);

            // 主标题：22pt 加粗 深蓝
            AddPara(rng, "一元二次方程讲义（TypeRight 示例）", r =>
            {
                r.Font.Size = 22;
                r.Font.Bold = -1;
                r.Font.Color = WdColor.wdColorDarkBlue;
            });

            AddPara(rng, "", null);

            // 模块标题 ×2：15pt 加粗 深蓝（与主标题仅字号不同 → 演示「忽略字号」）
            AddPara(rng, "模块一：一元二次方程的本质", r =>
            {
                r.Font.Size = 15;
                r.Font.Bold = -1;
                r.Font.Color = WdColor.wdColorDarkBlue;
            });
            AddPara(rng, "模块二：解法全景与选择", r =>
            {
                r.Font.Size = 15;
                r.Font.Bold = -1;
                r.Font.Color = WdColor.wdColorDarkBlue;
            });

            // 小节标题 ×3：12.5pt 加粗 中蓝
            AddPara(rng, "1.1 一元二次方程的标准定义", SectionFormat);
            AddPara(rng, "1.1.2 根的判别", SectionFormat);
            AddPara(rng, "1.2 四大解法剖析", SectionFormat);

            // 正文 ×4：10.5pt 无格式（其中混入纯文本公式与局部加粗）
            AddPara(rng, "对于方程 ax² + bx + c = 0，其中 a≠0，我们称 x₁ 和 x₂ 为它的两个根。", null);
            AddPara(rng, "只含一个未知数，且未知数的最高次数是 2 的整式方程叫作一元二次方程。", null);
            AddPara(rng, "将各项系数代入求根公式(-b±√(b²-4ac))/(2a)即可求出两根。", null);
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

            // 独立公式行（正文签名）
            AddPara(rng, "x₁² + x₂² = (x₁ + x₂)² - 2x₁x₂", null);

            // 图注：9pt 灰
            AddPara(rng, "图 1 一元二次函数的图像", r =>
            {
                r.Font.Size = 9;
                r.Font.Color = WdColor.wdColorGray50;
            });

            AddTable(doc, rng);

            AddPara(rng, "综上，理解判别式的几何意义，是掌握二次函数与一元二次方程内在联系的关键一步。", null);

            return doc;
        }

        private static void SectionFormat(Range r)
        {
            r.Font.Size = 12.5f;
            r.Font.Bold = -1;
            r.Font.Color = WdColor.wdColorBlue;
        }

        private static void AddPara(Range rng, string text, Action<Range> format)
        {
            rng.Collapse(WdCollapseDirection.wdCollapseEnd);
            rng.InsertAfter(text);
            rng.InsertParagraphAfter();
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
            table.Rows[1].Range.Font.Bold = -1;
            table.Cell(2, 1).Range.Text = "a";
            table.Cell(2, 2).Range.Text = "二次项系数，a≠0";
            table.Cell(3, 1).Range.Text = "Δ";
            table.Cell(3, 2).Range.Text = "b² - 4ac，决定根的个数";
        }
    }
}
