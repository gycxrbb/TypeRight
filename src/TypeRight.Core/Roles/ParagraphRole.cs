using System;
using System.Collections.Generic;

namespace TypeRight.Roles
{
    /// <summary>段落语义角色（PRD §4.1）。</summary>
    public enum ParagraphRole
    {
        Unknown = 0,
        Title,
        Heading1,
        Heading2,
        Heading3,
        Body,
        Caption,
        ListItem,
        TableText,
        TableHeader,
        FormulaLine,
        Skip
    }

    public static class ParagraphRoles
    {
        public const string KeyTitle = "title";
        public const string KeyHeading1 = "heading1";
        public const string KeyHeading2 = "heading2";
        public const string KeyHeading3 = "heading3";
        public const string KeyBody = "body";
        public const string KeyCaption = "caption";
        public const string KeyListItem = "listItem";
        public const string KeyTableText = "tableText";
        public const string KeyFormulaLine = "formulaLine";
        public const string KeySkip = "skip";

        private static readonly Dictionary<ParagraphRole, string> ToKeyMap = new Dictionary<ParagraphRole, string>
        {
            { ParagraphRole.Title, KeyTitle },
            { ParagraphRole.Heading1, KeyHeading1 },
            { ParagraphRole.Heading2, KeyHeading2 },
            { ParagraphRole.Heading3, KeyHeading3 },
            { ParagraphRole.Body, KeyBody },
            { ParagraphRole.Caption, KeyCaption },
            { ParagraphRole.ListItem, KeyListItem },
            { ParagraphRole.TableText, KeyTableText },
            { ParagraphRole.TableHeader, KeyTableText },
            { ParagraphRole.FormulaLine, KeyFormulaLine },
            { ParagraphRole.Skip, KeySkip }
        };

        private static readonly Dictionary<string, ParagraphRole> FromKeyMap = new Dictionary<string, ParagraphRole>
        {
            { KeyTitle, ParagraphRole.Title },
            { KeyHeading1, ParagraphRole.Heading1 },
            { KeyHeading2, ParagraphRole.Heading2 },
            { KeyHeading3, ParagraphRole.Heading3 },
            { KeyBody, ParagraphRole.Body },
            { KeyCaption, ParagraphRole.Caption },
            { KeyListItem, ParagraphRole.ListItem },
            { KeyTableText, ParagraphRole.TableText },
            { "tableHeader", ParagraphRole.TableHeader },
            { KeyFormulaLine, ParagraphRole.FormulaLine },
            { KeySkip, ParagraphRole.Skip }
        };

        /// <summary>角色在 Profile roles 字典中的键名。</summary>
        public static string ToProfileKey(this ParagraphRole role)
        {
            string key;
            return ToKeyMap.TryGetValue(role, out key) ? key : KeyBody;
        }

        /// <summary>解析 Profile / LLM 返回的键名，无法识别时返回 null。</summary>
        public static ParagraphRole? TryParse(string key)
        {
            if (string.IsNullOrEmpty(key)) return null;
            ParagraphRole role;
            key = key.Trim();
            return FromKeyMap.TryGetValue(key, out role) ? (ParagraphRole?)role : null;
        }

        /// <summary>中文标签（预览界面用）。</summary>
        public static string ToLabel(this ParagraphRole role)
        {
            switch (role)
            {
                case ParagraphRole.Title: return "文档主标题";
                case ParagraphRole.Heading1: return "一级标题";
                case ParagraphRole.Heading2: return "二级标题";
                case ParagraphRole.Heading3: return "三级标题";
                case ParagraphRole.Body: return "正文";
                case ParagraphRole.Caption: return "图注/表注";
                case ParagraphRole.ListItem: return "列表项";
                case ParagraphRole.TableText: return "表格文字";
                case ParagraphRole.TableHeader: return "表头";
                case ParagraphRole.FormulaLine: return "独立公式行";
                case ParagraphRole.Skip: return "跳过";
                default: return "未识别";
            }
        }

        public static bool IsHeading(this ParagraphRole role)
        {
            return role == ParagraphRole.Heading1 || role == ParagraphRole.Heading2 || role == ParagraphRole.Heading3;
        }

        /// <summary>TableHeader 在 Profile 中没有独立条目，格式落点等同 tableText（加粗保留）。</summary>
        public static string FormatKey(this ParagraphRole role)
        {
            return role == ParagraphRole.TableHeader ? KeyTableText : role.ToProfileKey();
        }
    }
}
