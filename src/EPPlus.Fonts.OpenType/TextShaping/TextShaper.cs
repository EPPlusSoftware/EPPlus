/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/15/2025         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    public class TextShaper
    {
        private readonly OpenTypeFont _font;
        private readonly KerningProvider _kerningProvider;
        private readonly LigatureProcessor _ligatureProcessor;

        public TextShaper(OpenTypeFont font)
        {
            _font = font ?? throw new ArgumentNullException(nameof(font));
            _kerningProvider = new KerningProvider(font);
            _ligatureProcessor = new LigatureProcessor(font);
        }

        #region Single-line Shaping

        /// <summary>
        /// Shape text using default options (ligatures + kerning).
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <returns>Shaped text with positioned glyphs</returns>
        public ShapedText Shape(string text)
        {
            return Shape(text, ShapingOptions.Default);
        }

        /// <summary>
        /// Shape text with specified options.
        /// Note: Newline characters (\n, \r, \r\n) are treated as regular characters.
        /// For multi-line text, use ShapeLines() method instead.
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <param name="options">Shaping options</param>
        /// <returns>Shaped text with positioned glyphs</returns>
        public ShapedText Shape(string text, ShapingOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedText
                {
                    OriginalText = text ?? string.Empty,
                    Glyphs = new ShapedGlyph[0]
                };
            }

            if (options == null)
            {
                options = ShapingOptions.Default;
            }

            // Phase 1: Map characters to glyphs
            var glyphs = MapToGlyphs(text);

            // Phase 2: Apply GSUB substitutions (if enabled)
            if (options.ApplySubstitutions && _font.GsubTable != null)
            {
                glyphs = ApplyGsubSubstitutions(glyphs, options);
            }

            // Phase 3: Apply GPOS positioning (if enabled)
            if (options.ApplyPositioning)
            {
                ApplyPositioning(glyphs, options);
            }

            // Phase 4: Build result
            return new ShapedText
            {
                OriginalText = text,
                Glyphs = glyphs.ToArray()
            };
        }

        #endregion

        #region Phase 1: Character to Glyph Mapping

        /// <summary>
        /// Maps characters to glyphs using the cmap table.
        /// </summary>
        private List<ShapedGlyph> MapToGlyphs(string text)
        {
            var glyphs = new List<ShapedGlyph>(text.Length);
            var cmapTable = _font.CmapTable;
            var hmtxTable = _font.HmtxTable;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                // Map character to glyph ID
                int glyphId = cmapTable.MapCharToGlyph(c);

                // Handle missing glyphs (use .notdef)
                if (glyphId < 0)
                {
                    glyphId = 0; // .notdef
                }

                // Get advance width from hmtx
                int advanceWidth = hmtxTable.GetAdvanceWidth((ushort)glyphId);

                glyphs.Add(new ShapedGlyph
                {
                    GlyphId = (ushort)glyphId,
                    XAdvance = advanceWidth,
                    YAdvance = 0,
                    XOffset = 0,
                    YOffset = 0,
                    ClusterIndex = i,
                    CharCount = 1
                });
            }

            return glyphs;
        }

        #endregion

        #region Phase 2: GSUB Substitutions

        /// <summary>
        /// Applies GSUB substitutions (ligatures, contextual alternates, etc.).
        /// </summary>
        private List<ShapedGlyph> ApplyGsubSubstitutions(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            // TODO: Implement in next step
            // For now, just apply ligatures if "liga" feature is requested

            if (options.GsubFeatures != null && options.GsubFeatures.Contains("liga"))
            {
                glyphs = _ligatureProcessor.ApplyLigatures(glyphs);
            }

            return glyphs;
        }

        #endregion

        #region Phase 3: Positioning

        /// <summary>
        /// Applies positioning adjustments (kerning, mark positioning, etc.).
        /// </summary>
        private void ApplyPositioning(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            // Apply kerning if requested
            if (options.GposFeatures != null && options.GposFeatures.Contains("kern"))
            {
                ApplyKerning(glyphs);
            }

            // TODO: Add support for other GPOS features (mark, mkmk, etc.)
        }

        /// <summary>
        /// Applies kerning adjustments to glyph pairs.
        /// Uses KerningProvider which handles both GPOS and legacy kern table.
        /// </summary>
        private void ApplyKerning(List<ShapedGlyph> glyphs)
        {
            for (int i = 1; i < glyphs.Count; i++)
            {
                ushort leftGlyph = glyphs[i - 1].GlyphId;
                ushort rightGlyph = glyphs[i].GlyphId;

                // Get kerning value (handles GPOS + kern table + caching)
                short kernValue = _kerningProvider.GetKerning(leftGlyph, rightGlyph);

                if (kernValue != 0)
                {
                    // Apply kerning to the left glyph's advance
                    var glyph = glyphs[i - 1];
                    glyph.XAdvance += kernValue;
                    glyphs[i - 1] = glyph;
                }
            }
        }

        #endregion

        #region Utilities

        /// <summary>
        /// Measures the width of text in font units.
        /// </summary>
        public int MeasureText(string text, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            return shaped.TotalAdvanceWidth;
        }

        /// <summary>
        /// Measures the width of text in PDF points.
        /// </summary>
        public float MeasureTextInPoints(string text, float fontSize, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;
            return shaped.GetWidthInPoints(fontSize, unitsPerEm);
        }

        /// <summary>
        /// Measures the width of text in pixels.
        /// </summary>
        public float MeasureTextInPixels(string text, float fontSize, float dpi, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;
            return shaped.GetWidthInPixels(fontSize, dpi, unitsPerEm);
        }

        #endregion

        #region Multi-line Support

        /// <summary>
        /// Shape multi-line text (handles \n, \r, \r\n).
        /// Returns one ShapedText per line.
        /// </summary>
        public ShapedText[] ShapeLines(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedText[0];
            }

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var result = new ShapedText[lines.Length];

            for (int i = 0; i < lines.Length; i++)
            {
                result[i] = Shape(lines[i], options);
            }

            return result;
        }

        /// <summary>
        /// Measure multi-line text and return bounding box.
        /// </summary>
        public MultiLineMetrics MeasureLines(string text, float fontSize, ShapingOptions options = null)
        {
            var shapedLines = ShapeLines(text, options);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;

            float maxWidth = 0;
            foreach (var line in shapedLines)
            {
                float lineWidth = line.GetWidthInPoints(fontSize, unitsPerEm);
                maxWidth = Math.Max(maxWidth, lineWidth);
            }

            float lineHeight = GetLineHeightInPoints(fontSize);
            float fontHeight = GetFontHeightInPoints(fontSize);
            float totalHeight = shapedLines.Length * lineHeight;

            return new MultiLineMetrics
            {
                Width = maxWidth,
                Height = totalHeight,
                FontHeight = fontHeight,
                LineCount = shapedLines.Length,
                LineHeight = lineHeight
            };
        }

        /// <summary>
        /// Get line height (ascent + descent + line gap) in points.
        /// </summary>
        public float GetLineHeightInPoints(float fontSize)
        {
            var hhea = _font.HheaTable;
            float unitsPerEm = _font.HeadTable.UnitsPerEm;

            // ascent is positive, descender is negative
            int lineHeightUnits = hhea.ascender - hhea.descender + hhea.lineGap;

            return (lineHeightUnits / unitsPerEm) * fontSize;
        }

        /// <summary>
        /// Get font height (ascent + descent only, no line gap) in points.
        /// </summary>
        public float GetFontHeightInPoints(float fontSize)
        {
            var hhea = _font.HheaTable;
            float unitsPerEm = _font.HeadTable.UnitsPerEm;

            // ascent is positive, descender is negative
            int fontHeightUnits = hhea.ascender - hhea.descender;

            return (fontHeightUnits / unitsPerEm) * fontSize;
        }

        #endregion
    }
}