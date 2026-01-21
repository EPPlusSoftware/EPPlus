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
        /// </summary>
        private List<string> WrapParagraph(
            string text,
            float fontSize,
            double maxWidthPoints,
            double startingWidthPoints,
            ShapingOptions options)
        {
            var lines = new List<string>();

            // Track current line being built
            var currentLine = new System.Text.StringBuilder();
            var currentWord = new System.Text.StringBuilder();

            double currentLineWidth = startingWidthPoints;
            double currentWordWidth = 0;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                if (c == ' ')
                {
                    // Try to add the word + space to current line
                    string wordText = currentWord.ToString();
                    double spaceWidth = MeasureText(" ", fontSize, options);
                    double totalWidth = currentLineWidth + currentWordWidth + spaceWidth;

                    if (totalWidth <= maxWidthPoints || currentLine.Length == 0)
                    {
                        // Word fits on current line
                        if (currentLine.Length > 0)
                        {
                            currentLine.Append(' ');
                            currentLineWidth += spaceWidth;
                        }
                        currentLine.Append(wordText);
                        currentLineWidth += currentWordWidth;

                        // Reset word (same as Clear(), works in NET35)
                        currentWord.Length = 0;
                        currentWordWidth = 0;
                    }
                    else
                    {
                        // Word doesn't fit - wrap to new line
                        lines.Add(currentLine.ToString());

                        // Reset line and start with word (same as Clear(), works in NET35)
                        currentLine.Length = 0;
                        currentLine.Append(wordText);
                        currentLineWidth = currentWordWidth;

                        // Reset word (same as Clear(), works in NET35)
                        currentWord.Length = 0;
                        currentWordWidth = 0;
                    }
                }
                else
                {
                    // Add character to current word
                    currentWord.Append(c);

                    // Measure word so far (with proper shaping)
                    string wordSoFar = currentWord.ToString();
                    currentWordWidth = MeasureText(wordSoFar, fontSize, options);

                    // Check if word itself is too long for a line
                    if (currentLineWidth + currentWordWidth > maxWidthPoints && currentLine.Length > 0)
                    {
                        // Wrap current line and start new line with this word
                        lines.Add(currentLine.ToString());
                        // same as Clear(), works in NET35
                        currentLine.Length = 0;
                        currentLineWidth = 0;
                    }
                }
            }

            // Add remaining word and line
            if (currentWord.Length > 0)
            {
                string wordText = currentWord.ToString();

                if (currentLine.Length > 0 && currentLineWidth + currentWordWidth > maxWidthPoints)
                {
                    // Word doesn't fit - wrap to new line
                    lines.Add(currentLine.ToString());
                    // same as Clear(), works in NET35
                    currentLine.Length = 0;
                    currentLine.Append(wordText);
                }
                else
                {
                    // Word fits
                    if (currentLine.Length > 0)
                    {
                        currentLine.Append(' ');
                    }
                    currentLine.Append(wordText);
                }
            }

            if (currentLine.Length > 0)
            {
                lines.Add(currentLine.ToString());
            }

            // Ensure at least one line
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