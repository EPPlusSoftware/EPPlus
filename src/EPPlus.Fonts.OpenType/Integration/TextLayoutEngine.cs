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
  01/24/2025         EPPlus Software AB           Added StringBuilder pooling (.NET 3.5 compatible)
  05/06/2026         EPPlus Software AB           Removed per-instance font directories — uses global config
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.Utilities;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
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

        // Space width cache - avoids repeated Shape(" ") calls
        private readonly Dictionary<float, double> _spaceWidthCache;

        // ArrayPool buffer - only ONE buffer for entire class
        private double[] _charWidthBuffer = null;
        private int _charWidthBufferCapacity = 0;

        // StringBuilder pooling - reuse between wrapping operations
        private readonly StringBuilder _lineBuilder = new StringBuilder(256);

        private List<string> _lineListBuffer = new List<string>(256);
        private bool _disposed = false;

        /// <summary>
        /// Creates a TextLayoutEngine for single-font text wrapping.
        /// Font resolution for rich-text fragments uses the globally configured resolver — to
        /// search additional directories or install a custom resolver, use
        /// <see cref="OpenTypeFonts.Configure"/>.
        /// </summary>
        public TextLayoutEngine(ITextShaper shaper)
        {
            _shaper = shaper ?? throw new ArgumentNullException(nameof(shaper));
            _spaceWidthCache = new Dictionary<float, double>();
        }

        public double GetLineHeightInPoints(float fontSize)
        {
            return _shaper.GetLineHeightInPoints(fontSize);
        }

        public double GetBaseLineInPoints(float fontSize)
        {
            return _shaper.GetAscentInPoints(fontSize);
        }

        public double GetDescentInPoints(float fontSize)
        {
            return _shaper.GetDescentInPoints(fontSize);
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

            //paragraphs that have endline symbols should keep trailing spaces
            //others should not. Add an extra space as the trailing space is always trimmed
            //if (paragraphs.Length > 1)
            //{
            //    for (int i = 1; i < paragraphs.Length - 1; i++)
            //    {
            //        paragraphs[i] = paragraphs[i] + " ";
            //    }
            //}

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
                return CreateEmptyResult();
            }

            var charWidths = CalculateCharacterWidths(text, fontSize, options);

            var state = new WrapStateText(startingWidthPoints, GetCachedSpaceWidth(fontSize, options));

            PrepareLineBuilder(text.Length);

            for (int i = 0; i <= text.Length; i++)
            {
                var charType = GetCharacterType(text, i);

                if (state.IsCompleteWordReady(charType, i))  // ← Använd state
                {
                    ProcessCompleteWord(text, state, i, maxWidthPoints);
                }
                else if (charType == CharacterType.Space)
                {
                    ProcessNonEndingSpace(text, state, i, maxWidthPoints);
                }
                else if (charType == CharacterType.Regular)
                {
                    ProcessCharacterInWord(text, charWidths, state, i, maxWidthPoints);
                }
                else
                {
                    if (text[i - 1] == ' ')
                    {   //Add extra to avoid trimming
                        _lineBuilder.Append("  ");
                        //if (_lineBuilder.LastChar() == ' ')
                        //{
                        //    //Add extra to avoid trimming
                        //    _lineBuilder.Append(" ");
                        //}
                    }
                }
            }

            return FinalizeWrapping();
        }


        private double MeasureText(string text, float fontSize, ShapingOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var shaped = _shaper.Shape(text, options);
            return shaped.GetWidthInPoints(fontSize);
        }

        private ITextShaper GetShaperForFont(MeasurementFont font)
        {
            return OpenTypeFonts.GetShaperForFont(font);
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

                    // Clear StringBuilder to release string references (.NET 3.5 compatible)
                    _lineBuilder.Length = 0;

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