using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace TypeRight.Formatting
{
    /// <summary>
    /// 示例文档生成器：故意用"脏"直接格式（字体混排、字号不一、纯文本公式、表格），
    /// 让用户在零风险环境里体验"扫描 → 确认 → 一键重排 → Ctrl+Z 撤销"的完整流程。
    /// </summary>
    public static class SampleDocBuilder
    {
        public static Document Build(Application app)
        {
            if (app == null) throw new ArgumentNullException("app");
            Document doc = app.Documents.Add();

            Range rng = doc.Range(0, 0);

            // 主标题：直接格式 22pt 加粗
            AddPara(rng, "一元二次方程讲义（TypeRight 示例文档）", r => { r.Font.Size = 22; r.Font.Bold = -1; });
            AddPara(rng, "", null);
            // 一级标题（纯文本无格式）
            AddPara(rng, "第一章 方程与函数", null);
            AddPara(rng, "1.1 一元二次方程的定义", null);

            // 正文：混排 + 纯文本公式
            AddPara(rng, "对于方程 ax² + bx + c = 0，其中 a≠0，我们称 x₁ 和 x₂ 为它的两个根。判别式 Δ = b² - 4ac 决定根的个数。", null);
            AddPara(rng, "只含一个未知数，且未知数的最高次数是 2 的整式方程叫作一元二次方程。标准形式中 a、b、c 分别称为二次项系数、一次项系数和常数项。", null);

            AddPara(rng, "1.1.2 根的判别", null);
            AddPara(rng, "将各项系数代入求根公式(-b±√(b²-4ac))/(2a)即可求出两根。", null);

            // 独立公式行
            AddPara(rng, "x₁² + x₂² = (x₁ + x₂)² - 2x₁x₂", null);

            AddPara(rng, "图 1 一元二次函数的图像", null);
            AddPara(rng, "模块二：例题精讲", null);

            // 局部加粗强调
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

            AddPara(rng, "一元二次方程的求根公式最早出现在公元 9 世纪波斯数学家花拉子米的《代数学》中，是中学代数的核心内容之一。", null);
            AddPara(rng, "表 1 系数含义", null);

            AddTable(doc, rng);

            AddPara(rng, "第二章 二次函数与方程的关系", null);
            AddPara(rng, "二次函数 y = ax² + bx + c 的图像与 x 轴的交点个数，恰好对应一元二次方程 ax² + bx + c = 0 的实根个数：Δ > 0 时有两个交点，Δ = 0 时有一个，Δ < 0 时没有交点。这就是数形结合的基本思想。", null);
            AddPara(rng, "综上，理解判别式的几何意义，是掌握二次函数与一元二次方程内在联系的关键一步。", null);

            return doc;
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
