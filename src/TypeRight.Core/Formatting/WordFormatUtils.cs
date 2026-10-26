using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Profiles;

namespace TypeRight.Formatting
{
    /// <summary>Word 值约定与格式参数映射（WdColor 是 0x00BBGGRR；9999999 = wdUndefined 混合值）。</summary>
    public static class WordFormatUtils
    {
        public const float UndefinedFloat = 9999999f;
        public const int UndefinedInt = 9999999;
        public const int WdColorAutomatic = -16777216; // 0xFF000000

        /// <summary>WdColor(0x00BBGGRR) → RGB(0xRRGGBB)；自动色/混合返回 -1。</summary>
        public static int WdColorToRgb(int wdColor)
        {
            if (wdColor == UndefinedInt || wdColor == WdColorAutomatic || wdColor < 0 || wdColor > 0xFFFFFF) return -1;
            int r = wdColor & 0xFF;
            int g = (wdColor >> 8) & 0xFF;
            int b = (wdColor >> 16) & 0xFF;
            return (r << 16) | (g << 8) | b;
        }

        /// <summary>RGB(0xRRGGBB) → WdColor。</summary>
        public static int RgbToWdColor(int rgb)
        {
            int r = (rgb >> 16) & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = rgb & 0xFF;
            return r | (g << 8) | (b << 16);
        }

        /// <summary>解析 "#RRGGBB"（必须带 #）；非法返回 null。</summary>
        public static int? ParseHexColor(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return null;
            hex = hex.Trim();
            if (hex.Length != 7 || hex[0] != '#') return null;
            int v;
            if (!int.TryParse(hex.Substring(1), System.Globalization.NumberStyles.HexNumber, null, out v)) return null;
            return v;
        }

        public static WdParagraphAlignment ToWdAlignment(ProfileAlignment alignment)
        {
            switch (alignment)
            {
                case ProfileAlignment.Left: return WdParagraphAlignment.wdAlignParagraphLeft;
                case ProfileAlignment.Center: return WdParagraphAlignment.wdAlignParagraphCenter;
                case ProfileAlignment.Right: return WdParagraphAlignment.wdAlignParagraphRight;
                case ProfileAlignment.Justify: return WdParagraphAlignment.wdAlignParagraphJustify;
                default: return WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        /// <summary>1.5 倍类行距 → Word 行距规则/值。多倍行距用 LineSpacing=12×倍数。</summary>
        public static void ApplyLineSpacing(ParagraphFormat format, double multiple)
        {
            if (format == null) return;
            double eps = 0.01;
            if (Math.Abs(multiple - 1.0) < eps)
            {
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
            else if (Math.Abs(multiple - 1.5) < eps)
            {
                format.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
            }
            else if (Math.Abs(multiple - 2.0) < eps)
            {
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
            }
            else
            {
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                format.LineSpacing = (float)(12.0 * multiple);
            }
        }

        public static WdOutlineLevel ToWdOutlineLevel(int level)
        {
            switch (level)
            {
                case 1: return WdOutlineLevel.wdOutlineLevel1;
                case 2: return WdOutlineLevel.wdOutlineLevel2;
                case 3: return WdOutlineLevel.wdOutlineLevel3;
                case 4: return WdOutlineLevel.wdOutlineLevel4;
                case 5: return WdOutlineLevel.wdOutlineLevel5;
                case 6: return WdOutlineLevel.wdOutlineLevel6;
                case 7: return WdOutlineLevel.wdOutlineLevel7;
                case 8: return WdOutlineLevel.wdOutlineLevel8;
                case 9: return WdOutlineLevel.wdOutlineLevel9;
                default: return WdOutlineLevel.wdOutlineLevelBodyText;
            }
        }

        /// <summary>内置样式名 → WdBuiltinStyle；非内置返回 null。</summary>
        public static WdBuiltinStyle? TryGetBuiltinStyle(string wordStyle)
        {
            if (string.IsNullOrEmpty(wordStyle)) return null;
            switch (wordStyle.Trim().ToLowerInvariant())
            {
                case "title": return WdBuiltinStyle.wdStyleTitle;
                case "heading 1": return WdBuiltinStyle.wdStyleHeading1;
                case "heading 2": return WdBuiltinStyle.wdStyleHeading2;
                case "heading 3": return WdBuiltinStyle.wdStyleHeading3;
                case "normal": return WdBuiltinStyle.wdStyleNormal;
                case "caption": return WdBuiltinStyle.wdStyleCaption;
                case "list paragraph": return WdBuiltinStyle.wdStyleListParagraph;
                default: return null;
            }
        }

        /// <summary>字号读取：9999999（混合）→ NaN。</summary>
        public static double NormalizeSize(float size)
        {
            return size >= UndefinedFloat ? double.NaN : size;
        }

        public static double NormalizeSize(int size)
        {
            return size >= UndefinedInt ? double.NaN : size;
        }

        // ---- 幂等写入助手：值相同则不写、不计数 ----

        /// <summary>读当前值，等于目标值则跳过写入。返回是否发生了改动。</summary>
        public static bool SetIfChanged<T>(Func<T> get, Action<T> set, T value, Func<T, T, bool> equals = null)
        {
            T current;
            try { current = get(); }
            catch { return false; }
            bool same = equals != null ? equals(current, value) : Equals(current, value);
            if (same) return false;
            try { set(value); }
            catch { return false; }
            return true;
        }

        public static bool FloatEq(float a, float b) { return Math.Abs(a - b) < 0.05f; }

        public static bool IntEq(int a, int b) { return a == b; }
    }
}
