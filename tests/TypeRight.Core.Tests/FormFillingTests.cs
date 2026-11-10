using System;
using System.Drawing;
using TypeRight.Forms;
using Xunit;

namespace TypeRight.Tests
{
    public class FormFillingTests
    {
        [Theory]
        [InlineData("申请人（手签）：", true)]
        [InlineData("学院主管学生工作领导签字（盖章）：", true)]
        [InlineData("盖章：", true)]
        [InlineData("2026 年 11 月 9 日", true)]
        [InlineData("   年   月   日  \r\a", true)]
        [InlineData("本人思想端正，学习刻苦，成绩优异。", false)]
        [InlineData("", false)]
        public void IsSignatureOrDateLine_DetectsBottomReserveLines(string text, bool expected)
        {
            Assert.Equal(expected, FormGeometry.IsSignatureOrDateLine(text));
        }

        [Fact]
        public void CleanCellText_StripsWordCellEndMarkers()
        {
            string raw = "首都经济贸易大学\r\a";
            Assert.Equal("首都经济贸易大学", FormGeometry.CleanCellText(raw));
        }

        [Fact]
        public void CropBitmapToAspect_PreservesTargetAspectRatio_WithoutDistortion()
        {
            using (var src = new Bitmap(300, 500)) // 0.60 偏瘦高图片
            {
                float targetAspect = 0.75f; // 3:4 单元格
                using (Bitmap cropped = PhotoCellFitter.CropBitmapToAspect(src, targetAspect))
                {
                    Assert.Equal(300, cropped.Width);
                    Assert.Equal(400, cropped.Height); // 300 / 0.75 = 400
                }
            }

            using (var wideSrc = new Bitmap(500, 400)) // 1.25 偏宽图片
            {
                float targetAspect = 0.75f;
                using (Bitmap cropped = PhotoCellFitter.CropBitmapToAspect(wideSrc, targetAspect))
                {
                    Assert.Equal(300, cropped.Width); // 400 * 0.75 = 300
                    Assert.Equal(400, cropped.Height);
                }
            }
        }

        [Fact]
        public void TextHalfWidth_CountsChineseAsTwoAndAsciiAsOne()
        {
            Assert.Equal(8, UnderlineFiller.TextHalfWidth("经济学院"));
            Assert.Equal(10, UnderlineFiller.TextHalfWidth("2025010203"));
            Assert.Equal(6, UnderlineFiller.TextHalfWidth("AI学院"));
        }
    }
}
