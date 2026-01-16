/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// ITextMeasurer implementation using OpenType font shaping.
    /// Provides accurate text measurement with ligatures and kerning support.
    /// </summary>
    public class OpenTypeFontTextMeasurer : ITextMeasurer
    {
        private readonly TextShaper _shaper;
        private readonly OpenTypeFont _font;
        private ShapingOptions _shapingOptions;

        public OpenTypeFontTextMeasurer(OpenTypeFont font, ShapingOptions options = null)
        {
            _font = font ?? throw new ArgumentNullException(nameof(font));
            _shaper = new TextShaper(font);
            _shapingOptions = options ?? ShapingOptions.Default;
            MeasureWrappedTextCells = true;
        }

        /// <summary>
        /// Always valid - pure .NET implementation with no external dependencies.
        /// </summary>
        public bool ValidForEnvironment()
        {
            return true;
        }

        /// <summary>
        /// Controls whether multi-line text (with CR/LF/CRLF) should be measured.
        /// </summary>
        public bool MeasureWrappedTextCells { get; set; }

        /// <summary>
        /// Measures text width and height.
        /// </summary>
        public TextMeasurement MeasureText(string text, MeasurementFont font)
        {
            if (string.IsNullOrEmpty(text))
            {
                return TextMeasurement.Empty;
            }

            // Check if text contains newlines
            bool hasNewlines = text.IndexOfAny(new[] { '\r', '\n' }) >= 0;

            if (hasNewlines && MeasureWrappedTextCells)
            {
                return MeasureMultiLineText(text, font);
            }
            else
            {
                return MeasureSingleLineText(text, font);
            }
        }

        private TextMeasurement MeasureSingleLineText(string text, MeasurementFont font)
        {
            var shaped = _shaper.Shape(text, _shapingOptions);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;

            float width = shaped.GetWidthInPoints(font.Size, unitsPerEm);
            float lineHeight = _shaper.GetLineHeightInPoints(font.Size);
            float fontHeight = _shaper.GetFontHeightInPoints(font.Size);

            return new TextMeasurement(width, lineHeight)
            {
                FontHeight = fontHeight
            };
        }

        private TextMeasurement MeasureMultiLineText(string text, MeasurementFont font)
        {
            var metrics = _shaper.MeasureLines(text, font.Size, _shapingOptions);

            return new TextMeasurement(metrics.Width, metrics.Height)
            {
                FontHeight = metrics.FontHeight
            };
        }
    }
}