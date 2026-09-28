/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/24/2026         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    /// <summary>
    /// Measurement helpers built purely on top of the <see cref="ITextShaper"/> contract
    /// (Shape/ShapeLines/GetLineHeightInPoints/GetFontHeightInPoints). Kept as extension
    /// methods so every ITextShaper implementation (TextShaper, GenericFontTextShaper, ...)
    /// gets them for free without duplicating the width-extraction logic and without adding
    /// members to the ITextShaper interface itself.
    /// </summary>
    public static class TextShaperExtensions
    {
        /// <summary>
        /// Measures the width of text in PDF points.
        /// </summary>
        public static float MeasureTextInPoints(this ITextShaper shaper, string text, float fontSize, ShapingOptions options = null)
        {
            var shaped = shaper.Shape(text, options);
            return shaped.GetWidthInPoints(fontSize);
        }

        /// <summary>
        /// Measures the width of text in pixels.
        /// </summary>
        public static float MeasureTextInPixels(this ITextShaper shaper, string text, float fontSize, float dpi, ShapingOptions options = null)
        {
            var shaped = shaper.Shape(text, options);
            return shaped.GetWidthInPixels(fontSize, dpi);
        }

        /// <summary>
        /// Measures multi-line text and returns its bounding box.
        /// </summary>
        public static MultiLineMetrics MeasureLines(this ITextShaper shaper, string text, float fontSize, ShapingOptions options = null)
        {
            var shapedLines = shaper.ShapeLines(text, options);

            float maxWidth = 0;
            for (var i = 0; i < shapedLines.Length; i++)
            {
                var lineWidth = shapedLines[i].GetWidthInPoints(fontSize);
                if (lineWidth > maxWidth)
                {
                    maxWidth = lineWidth;
                }
            }

            var lineHeight = shaper.GetLineHeightInPoints(fontSize);
            var fontHeight = shaper.GetFontHeightInPoints(fontSize);

            return new MultiLineMetrics
            {
                Width = maxWidth,
                Height = shapedLines.Length * lineHeight,
                FontHeight = fontHeight,
                LineCount = shapedLines.Length,
                LineHeight = lineHeight
            };
        }
    }
}