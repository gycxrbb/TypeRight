using System;
using TypeRight.Roles;

namespace TypeRight.Scanning
{
    /// <summary>
    /// 段落特征向量（PRD §4.2 第一段）。由 DocumentScanner 从 Word 对象模型提取，
    /// 角色判定只依赖此结构，保证规则引擎可单测。
    /// </summary>
    public sealed class ParagraphFeatures
    {
        public int Index;

        /// <summary>去除段落标记后的可见文本（\r、\a 已剔除）。</summary>
        public string Text = "";

        public bool IsEmpty;

        // ---- 字体特征（首字符所在 run）----
        /// <summary>字号 pt；NaN 表示混合/未定义（9999999）。</summary>
        public double FontSize = double.NaN;
        /// <summary>-1/0/9999999 → true/false/混合。</summary>
        public int BoldRaw;
        public bool Bold { get { return BoldRaw == -1; } }
        public bool BoldMixed { get { return BoldRaw == 9999999; } }
        /// <summary>字色 RGB（0xRRGGBB）；-1 表示混合/自动色。</summary>
        public int ColorRgb = -1;

        /// <summary>段落对齐："l"/"c"/"r"/"j"/"?"（左/中/右/两端/其他）。</summary>
        public string Alignment = "?";

        // ---- 结构特征 ----
        public bool InTable;
        public bool IsFirstTableRow;
        /// <summary>Word 自动编号/项目符号（ListFormat）。</summary>
        public bool HasAutoNumbering;
        /// <summary>已有 Heading 样式级别（0=无，1..3）。</summary>
        public int HeadingStyleLevel;
        public bool HasInlineImage;
        public bool HasFields;
        public bool IsTocField;

        // ---- 模式特征 ----
        public int NumberingPattern;   // 0 无；1 章（第X章/模块X）；2 两级数字；3 三级数字；4 中文顿号；5 带圈数字；6 【】
        public bool CaptionPattern;
        public double MathDensity;
        public bool HasEquals;
        public bool StandaloneMath;

        // ---- 分类阶段填充 ----
        public bool IsFirstNonEmpty;
        public ParagraphRole Role;
        public int RulePriority;             // 命中的规则编号（R1..R12）
        public bool HighConfidence;
        public string ClusterSignature;

        /// <summary>展示用截断文本。</summary>
        public string DisplayText(int max)
        {
            string t = Text;
            if (t.Length > max) t = t.Substring(0, max) + "…";
            return t;
        }

        public override string ToString()
        {
            return string.Format("#{0} [{1}] {2}", Index, Role, DisplayText(40));
        }
    }
}
