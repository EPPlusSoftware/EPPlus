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
  01/23/2025         EPPlus Software AB           Added ArrayPool optimization
  01/23/2025         EPPlus Software AB           Added space width cache
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.Utilities;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Handles text wrapping and layout using proper OpenType shaping.
    /// Replaces the old TextData wrapping logic.
    /// </summary>
    public partial class TextLayoutEngine : IDisposable
    {
        private readonly ITextShaper _shaper;
        private readonly List<string> _fontDirectories;
        private readonly bool _searchSystemDirectories;
        private readonly Dictionary<string, ITextShaper> _shaperCache;

        // Space width cache - avoids repeated Shape(" ") calls
        private readonly Dictionary<float, double> _spaceWidthCache;

        // ArrayPool buffer - endast EN buffer för hela klassen
        private double[] _charWidthBuffer = null;
        private int _charWidthBufferCapacity = 0;

        private List<string> _lineListBuffer = new List<string>(256);
        private bool _disposed = false;

        /// <summary>
        /// Creates a TextLayoutEngine for single-font text wrapping.
        /// </summary>
        public TextLayoutEngine(
            ITextShaper shaper,
            List<string> fontDirectories = null,
            bool searchSystemDirectories = true)
        {
            _shaper = shaper ?? throw new ArgumentNullException(nameof(shaper));
            _fontDirectories = fontDirectories ?? new List<string>();
            _searchSystemDirectories = searchSystemDirectories;
            _shaperCache = new Dictionary<string, ITextShaper>();
            _spaceWidthCache = new Dictionary<float, double>();
        }

        /// <summary>
        /// Gets a char width buffer with at least the specified capacity.
        /// Reuses existing buffer if large enough, otherwise rents larger one from pool.
        /// </summary>
        private double[] GetCharWidthBuffer(int minimumLength)
        {
            return ArrayPoolHelper<double>.EnsureCapacity(
                ref _charWidthBuffer,
                ref _charWidthBufferCapacity,
                minimumLength,
                clearArray: false
            );
        }

        /// <summary>
        /// Gets cached space width for the given font size.
        /// Caches the result to avoid repeated Shape(" ") calls.
        /// </summary>
        private double GetCachedSpaceWidth(float fontSize, ShapingOptions options)
        {
            // Check cache first
            if (_spaceWidthCache.TryGetValue(fontSize, out double cachedWidth))
            {
                return cachedWidth;
            }

            // Measure and cache
            double width = MeasureText(" ", fontSize, options);
            _spaceWidthCache[fontSize] = width;

            return width;
        }

        public List<string> WrapText(
            string text,
            float fontSize,
            double maxWidthPoints,
            ShapingOptions options = null)
        {
            return WrapText(text, fontSize, maxWidthPoints, 0, options);
        }

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

                double startingWidth = isFirstLine ? preExistingWidthPoints : 0;
                var wrappedLines = WrapParagraph(paragraph, fontSize, maxWidthPoints, startingWidth, options);
                lines.AddRange(wrappedLines);

                isFirstLine = false;
            }

            return lines;
        }

        private List<string> WrapParagraph(
            string text,
            float fontSize,
            double maxWidthPoints,
            double startingWidthPoints,
            ShapingOptions options)
        {
            _lineListBuffer.Clear();

            if (string.IsNullOrEmpty(text))
            {
                _lineListBuffer.Add(string.Empty);
                return new List<string>(_lineListBuffer);
            }

            // Get buffer from pool and extract widths
            var charWidths = GetCharWidthBuffer(text.Length);
            _shaper.ExtractCharWidths(text, fontSize, options, charWidths);

            // Use cached space width instead of measuring every time
            double spaceWidth = GetCachedSpaceWidth(fontSize, options);

            int lineStart = 0;
            int wordStart = 0;
            double currentLineWidth = startingWidthPoints;
            double currentWordWidth = 0;

            var currentLineBuilder = new StringBuilder(text.Length / 4 + 20);

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
                        if (currentLineBuilder.Length > 0)
                        {
                            currentLineBuilder.Append(' ');
                        }
                        currentLineBuilder.Append(text, wordStart, i - wordStart);
                        currentLineWidth = totalWidth;
                    }
                    else
                    {
                        if (currentLineBuilder.Length > 0 && currentLineBuilder[currentLineBuilder.Length - 1] == ' ')
                        {
                            currentLineBuilder.Length--;
                        }
                        if (currentLineBuilder.Length > 0)
                        {
                            _lineListBuffer.Add(currentLineBuilder.ToString());
                        }
                        currentLineBuilder.Length = 0;

                        lineStart = wordStart;
                        currentLineWidth = currentWordWidth;

                        currentLineBuilder.Append(text, wordStart, i - wordStart);
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
                        if (currentLineBuilder.Length > 0 && currentLineBuilder[currentLineBuilder.Length - 1] == ' ')
                        {
                            currentLineBuilder.Length--;
                        }
                        if (currentLineBuilder.Length > 0)
                        {
                            _lineListBuffer.Add(currentLineBuilder.ToString());
                        }
                        currentLineBuilder.Length = 0;

                        lineStart = wordStart;
                        currentLineWidth = 0;
                    }
                }
            }

            if (lineStart < text.Length)
            {
                if (currentLineBuilder.Length > 0 && currentLineBuilder[currentLineBuilder.Length - 1] == ' ')
                {
                    currentLineBuilder.Length--;
                }
                if (currentLineBuilder.Length > 0)
                {
                    _lineListBuffer.Add(currentLineBuilder.ToString());
                }
            }

            if (_lineListBuffer.Count == 0)
            {
                _lineListBuffer.Add(string.Empty);
            }

            return new List<string>(_lineListBuffer);
        }

        private double MeasureText(string text, float fontSize, ShapingOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var shaped = _shaper.Shape(text, options);
            return shaped.GetWidthInPoints(fontSize, _shaper.UnitsPerEm);
        }

        private double MeasureTextWithFont(string text, MeasurementFont font, ShapingOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var shaper = GetShaperForFont(font);
            var shaped = shaper.Shape(text, options ?? ShapingOptions.Default);
            return shaped.GetWidthInPoints(font.Size, shaper.UnitsPerEm);
        }

        private ITextShaper GetShaperForFont(MeasurementFont font)
        {
            string cacheKey = string.Format("{0}_{1}", font.FontFamily, GetFontSubFamily(font.Style));

            if (_shaperCache.TryGetValue(cacheKey, out var cachedShaper))
            {
                return cachedShaper;
            }

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

        #region IDisposable Implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Return buffer to pool
                    if (_charWidthBuffer != null)
                    {
                        ArrayPoolHelper<double>.SafeReturn(ref _charWidthBuffer, clearArray: false);
                        _charWidthBufferCapacity = 0;
                    }

                    // Dispose cached shapers
                    foreach (var shaper in _shaperCache.Values)
                    {
                        if (shaper is IDisposable disposable)
                        {
                            disposable.Dispose();
                        }
                    }
                    _shaperCache.Clear();

                    // Clear space width cache
                    _spaceWidthCache.Clear();
                }

                _disposed = true;
            }
        }

        ~TextLayoutEngine()
        {
            Dispose(false);
        }

        #endregion
    }
}