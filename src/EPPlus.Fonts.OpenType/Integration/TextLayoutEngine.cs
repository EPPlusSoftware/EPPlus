/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           TextLayoutEngine implementation
  01/22/2025         EPPlus Software AB           Optimized with shaping cache
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Handles text wrapping and layout using proper OpenType shaping.
    /// Replaces the old TextData wrapping logic.
    /// </summary>
    public partial class TextLayoutEngine
    {
        private readonly ITextShaper _shaper;
        private readonly List<string> _fontDirectories;
        private readonly bool _searchSystemDirectories;
        private readonly Dictionary<string, ITextShaper> _shaperCache;

        /// <summary>
        /// Creates a TextLayoutEngine for single-font text wrapping.
        /// </summary>
        /// <param name="shaper">Text shaper for the primary font</param>
        /// <param name="measurer">Text measurer</param>
        /// <param name="fontDirectories">Additional font directories to search (optional)</param>
        /// <param name="searchSystemDirectories">Whether to search system font directories</param>
        public TextLayoutEngine(
            ITextShaper shaper,
            List<string> fontDirectories = null,
            bool searchSystemDirectories = true)
        {
            _shaper = shaper ?? throw new ArgumentNullException(nameof(shaper));
            _fontDirectories = fontDirectories ?? new List<string>();
            _searchSystemDirectories = searchSystemDirectories;
            _shaperCache = new Dictionary<string, ITextShaper>();
        }

        /// <summary>
        /// Wraps text to fit within specified width.
        /// Handles word breaking at spaces and preserves existing line breaks.
        /// </summary>
        /// <param name="text">Text to wrap</param>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="maxWidthPoints">Maximum line width in points</param>
        /// <param name="options">Shaping options (null = default)</param>
        /// <returns>List of wrapped lines</returns>
        public List<string> WrapText(
            string text,
            float fontSize,
            double maxWidthPoints,
            ShapingOptions options = null)
        {
            return WrapText(text, fontSize, maxWidthPoints, 0, options);
        }

        /// <summary>
        /// Wraps text to fit within specified width with pre-existing content on first line.
        /// Used when text continues from previous content (e.g., different font on same line).
        /// </summary>
        /// <param name="text">Text to wrap</param>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="maxWidthPoints">Maximum line width in points</param>
        /// <param name="preExistingWidthPoints">Width already used on first line in points</param>
        /// <param name="options">Shaping options (null = default)</param>
        /// <returns>List of wrapped lines</returns>
        public List<string> WrapText(
            string text,
            float fontSize,
            double maxWidthPoints,
            double preExistingWidthPoints,
            ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new List<string> { string.Empty };
            }

            options = options ?? ShapingOptions.Default;
            var lines = new List<string>();

            // Handle existing line breaks first
            var paragraphs = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            bool isFirstLine = true;
            foreach (var paragraph in paragraphs)
            {
                if (string.IsNullOrEmpty(paragraph))
                {
                    lines.Add(string.Empty);
                    isFirstLine = false;
                    continue;
                }

                // Wrap this paragraph
                double startingWidth = isFirstLine ? preExistingWidthPoints : 0;
                var wrappedLines = WrapParagraph(paragraph, fontSize, maxWidthPoints, startingWidth, options);
                lines.AddRange(wrappedLines);

                isFirstLine = false;
            }

            return lines;
        }

        /// <summary>
        /// Wraps a single paragraph (no line breaks).
        /// OPTIMIZED: Minimizes memory allocations while maintaining O(n) performance.
        /// </summary>
        private List<string> WrapParagraph(
            string text,
            float fontSize,
            double maxWidthPoints,
            double startingWidthPoints,
            ShapingOptions options)
        {
            var lines = new List<string>();

            if (string.IsNullOrEmpty(text))
            {
                lines.Add(string.Empty);
                return lines;
            }

            // Shape once and build character width cache
            var charWidths = _shaper.ExtractCharWidths(text, fontSize, options);
            double spaceWidth = MeasureText(" ", fontSize, options);

            // Track word boundaries using indices
            int lineStart = 0;
            int wordStart = 0;
            double currentLineWidth = startingWidthPoints;
            double currentWordWidth = 0;

            for (int i = 0; i <= text.Length; i++)
            {
                bool isSpace = (i < text.Length && text[i] == ' ');
                bool isEnd = (i == text.Length);

                if ((isSpace || isEnd) && wordStart < i)
                {
                    double totalWidth = currentLineWidth + currentWordWidth;
                    if (lineStart < wordStart)
                    {
                        totalWidth += spaceWidth;
                    }

                    if (totalWidth <= maxWidthPoints || lineStart == wordStart)
                    {
                        // Word fits on current line
                        currentLineWidth = totalWidth;
                    }
                    else
                    {
                        // Word doesn't fit - wrap
                        lines.Add(text.Substring(lineStart, wordStart - lineStart).TrimEnd());
                        lineStart = wordStart;
                        currentLineWidth = currentWordWidth;
                    }

                    wordStart = i + 1;
                    currentWordWidth = 0;
                }
                else if (isSpace)
                {
                    wordStart = i + 1;
                }
                else
                {
                    currentWordWidth += charWidths[i];

                    if (currentWordWidth > maxWidthPoints && lineStart < wordStart && currentLineWidth > 0)
                    {
                        lines.Add(text.Substring(lineStart, wordStart - lineStart).TrimEnd());
                        lineStart = wordStart;
                        currentLineWidth = 0;
                    }
                }
            }

            // Add final line
            if (lineStart < text.Length)
            {
                lines.Add(text.Substring(lineStart).TrimEnd());
            }

            if (lines.Count == 0)
            {
                lines.Add(string.Empty);
            }

            return lines;
        }


        /// <summary>
        /// Measures text width using the primary shaper.
        /// </summary>
        private double MeasureText(string text, float fontSize, ShapingOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var shaped = _shaper.Shape(text, options);
            return shaped.GetWidthInPoints(fontSize, _shaper.UnitsPerEm);
        }

        /// <summary>
        /// Measures text width with a specific font (used for rich text).
        /// </summary>
        private double MeasureTextWithFont(string text, MeasurementFont font, ShapingOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            // Get or create shaper for this font
            var shaper = GetShaperForFont(font);

            // Shape and measure
            var shaped = shaper.Shape(text, options ?? ShapingOptions.Default);
            return shaped.GetWidthInPoints(font.Size, shaper.UnitsPerEm);
        }

        /// <summary>
        /// Gets or creates a TextShaper for the specified font.
        /// Uses caching to avoid creating multiple shapers for the same font.
        /// </summary>
        private ITextShaper GetShaperForFont(MeasurementFont font)
        {
            // Create cache key
            string cacheKey = $"{font.FontFamily}_{GetFontSubFamily(font.Style)}";

            // Check cache
            if (_shaperCache.TryGetValue(cacheKey, out var cachedShaper))
            {
                return cachedShaper;
            }

            // Load font and create shaper
            var openTypeFont = OpenTypeFonts.GetFontData(
                fontDirectories: _fontDirectories,
                fontName: font.FontFamily,
                subFamily: GetFontSubFamily(font.Style),
                searchSystemDirectories: _searchSystemDirectories
            );

            var shaper = new TextShaper(openTypeFont);
            _shaperCache[cacheKey] = shaper;

            return shaper;
        }

        /// <summary>
        /// Converts MeasurementFontStyles to FontSubFamily.
        /// </summary>
        private FontSubFamily GetFontSubFamily(MeasurementFontStyles style)
        {
            if ((style & (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic)) ==
                (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic))
            {
                return FontSubFamily.BoldItalic;
            }
            else if ((style & MeasurementFontStyles.Bold) == MeasurementFontStyles.Bold)
            {
                return FontSubFamily.Bold;
            }
            else if ((style & MeasurementFontStyles.Italic) == MeasurementFontStyles.Italic)
            {
                return FontSubFamily.Italic;
            }

            return FontSubFamily.Regular;
        }
    }
}