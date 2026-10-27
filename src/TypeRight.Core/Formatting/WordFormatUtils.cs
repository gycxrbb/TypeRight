using System;
using NetOffice.WordApi;

namespace TypeRight.Formatting
{
    /// <summary>Word 值约定（WdColor 是 0x00BBGGRR；9999999 = wdUndefined 混合值）。</summary>
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

        /// <summary>字号读取：9999999（混合）→ NaN。</summary>
        public static double NormalizeSize(float size)
        {
            return size >= UndefinedFloat ? double.NaN : size;
        }

        public static double NormalizeSize(int size)
        {
            return size >= UndefinedInt ? double.NaN : size;
        }

        /// <summary>读当前值，等于目标值则跳过写入（幂等口径）。返回是否发生了改动。</summary>
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
