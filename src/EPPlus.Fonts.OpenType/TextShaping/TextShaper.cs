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
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    /// <summary>
    /// Text shaping engine that converts text strings to positioned glyphs.
    /// Handles character-to-glyph mapping, GSUB substitutions, and GPOS positioning.
    /// </summary>
    public class TextShaper
    {
        private readonly OpenTypeFont _font;
        private readonly KerningProvider _kerningProvider;

        /// <summary>
        /// Creates a new text shaper for the specified font.
        /// </summary>
        /// <param name="font">The OpenType font to use for shaping</param>
        public TextShaper(OpenTypeFont font)
        {
            _font = font ?? throw new ArgumentNullException(nameof(font));
            _kerningProvider = new KerningProvider(font);
        }

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
                glyphs = ApplyLigatures(glyphs);
            }

            return glyphs;
        }

        /// <summary>
        /// Applies standard ligature substitutions (fi, ff, ffi, etc.).
        /// </summary>
        private List<ShapedGlyph> ApplyLigatures(List<ShapedGlyph> glyphs)
        {
            // TODO: Implement ligature lookup
            // For now, return unchanged
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
        /// <param name="text">Text to measure</param>
        /// <param name="options">Shaping options</param>
        /// <returns>Width in font units</returns>
        public int MeasureText(string text, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            return shaped.TotalAdvanceWidth;
        }

        /// <summary>
        /// Measures the width of text in PDF points.
        /// </summary>
        /// <param name="text">Text to measure</param>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="options">Shaping options</param>
        /// <returns>Width in PDF points</returns>
        public float MeasureTextInPoints(string text, float fontSize, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;
            return shaped.GetWidthInPoints(fontSize, unitsPerEm);
        }

        /// <summary>
        /// Measures the width of text in pixels.
        /// </summary>
        /// <param name="text">Text to measure</param>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="dpi">Screen DPI (typically 96)</param>
        /// <param name="options">Shaping options</param>
        /// <returns>Width in pixels</returns>
        public float MeasureTextInPixels(string text, float fontSize, float dpi, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;
            return shaped.GetWidthInPixels(fontSize, dpi, unitsPerEm);
        }

        #endregion
    }
}