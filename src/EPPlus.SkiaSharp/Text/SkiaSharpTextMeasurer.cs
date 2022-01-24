using OfficeOpenXml.Interfaces.Text;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.SkiaSharp.Text
{
    /// <summary>
    /// This class implements the <see cref="ITextMeasurer"/> interface
    /// and uses SkiaSharp to measure text.
    /// </summary>
    public class SkiaSharpTextMeasurer : ITextMeasurer
    {

        private SKFontStyle ToSkFontStyle(FontStyles style)
        {
            switch(style)
            {
                case FontStyles.Regular:
                    return SKFontStyle.Normal;
                case FontStyles.Bold:
                    return SKFontStyle.Bold;
                case FontStyles.Italic:
                    return SKFontStyle.Italic;
                case FontStyles.Bold | FontStyles.Italic:
                    return SKFontStyle.BoldItalic;
                default:
                    return SKFontStyle.Normal;
            }
        }

        /// <summary>
        /// Measures width and height of a text.
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="font"><see cref="ExcelFont">Font</see> of the measured text</param>
        /// <returns>A <see cref="TextMeasurement"/> instance with the width and the height in pixels</returns>
        public TextMeasurement MeasureText(string text, ExcelFont font)
        {
            var skFontStyle = ToSkFontStyle(font.Style);
            var tf = SKTypeface.FromFamilyName(font.FontFamily, skFontStyle);
            using(var paint = new SKPaint())
            {
                paint.TextSize = font.Size;
                paint.Typeface = tf;
                var rect = SKRect.Empty;
                paint.MeasureText(text.AsSpan(), ref rect);
                return new TextMeasurement(rect.Width * (96F / 72F), rect.Height * (96F / 72F));
            }
        }
    }
}
