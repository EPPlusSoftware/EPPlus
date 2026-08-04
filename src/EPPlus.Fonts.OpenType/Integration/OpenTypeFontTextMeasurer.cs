/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           OpenTypeFontTextMeasurer implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// ITextMeasurer implementation using OpenType font shaping.
    /// Provides accurate text measurement with ligatures and kerning support.
    /// </summary>
    public class OpenTypeFontTextMeasurer : ITextMeasurer
    {
        private readonly ITextShaper _shaper;
        private ShapingOptions _shapingOptions;

        public OpenTypeFontTextMeasurer(ITextShaper shaper, ShapingOptions options = null)
        {
            _shaper = shaper ?? throw new ArgumentNullException(nameof(shaper));
            _shapingOptions = options ?? ShapingOptions.Default;
            MeasureWrappedTextCells = true;
        }

        /// <summary>
        /// Always valid - pure .NET implementation with no external dependencies.
        /// </summary>
        public bool ValidForEnvironment() => true;

        /// <summary>
        /// Controls whether multi-line text (with CR/LF/CRLF) should be measured.
        /// </summary>
        public bool MeasureWrappedTextCells { get; set; }
        public eWrappedTextAutofitMode WrappedTextAutofitMode 
        { 
            get; 
            set; 
        }

        /// <summary>
        /// Measures text width and height.
        /// </summary>
        public TextMeasurement MeasureText(string text, MeasurementFont font)
        {
            if (string.IsNullOrEmpty(text))
            {
                // Return 0x0 for empty string, not TextMeasurement.Empty (-1x-1)
                return new TextMeasurement(0, 0);
            }


            // Check if text contains newlines
            bool hasNewlines = text.IndexOfAny(new[] { '\r', '\n' }) >= 0;

            if (hasNewlines && MeasureWrappedTextCells)
            {
                return MeasureMultiLineText(text, font.Size);
            }
            else
            {
                return MeasureSingleLineText(text, font.Size);
            }
        }

        /// <summary>
        /// Measures a single line of text.
        /// </summary>
        private TextMeasurement MeasureSingleLineText(string text, float fontSize)
        {
            var shaped = _shaper.Shape(text, _shapingOptions);

            // Convert from design units to points
            float width = shaped.GetWidthInPoints(fontSize);
            float lineHeight = (float)_shaper.GetLineHeightInPoints(fontSize);
            float fontHeight = (float)_shaper.GetFontHeightInPoints(fontSize);

            return new TextMeasurement(width, lineHeight)
            {
                FontHeight = fontHeight
            };
        }

        /// <summary>
        /// Measures multiple lines of text (separated by CR/LF/CRLF).
        /// Returns the maximum width and total height.
        /// </summary>
        private TextMeasurement MeasureMultiLineText(string text, float fontSize)
        {
            var shapedLines = _shaper.ShapeLines(text, _shapingOptions);

            // Calculate max width across all lines
            float maxWidth = 0;
            foreach (var line in shapedLines)
            {
                float lineWidth = line.GetWidthInPoints(fontSize);
                maxWidth = Math.Max(maxWidth, lineWidth);
            }

            float lineHeight = (float)_shaper.GetLineHeightInPoints(fontSize);
            float fontHeight = (float)_shaper.GetFontHeightInPoints(fontSize);
            float totalHeight = shapedLines.Length * lineHeight;

            return new TextMeasurement(maxWidth, totalHeight)
            {
                FontHeight = fontHeight
            };
        }
    }
}