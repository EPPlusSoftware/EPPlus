/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           Vertical text shaping support (Excel rotation 255)
                                                  Requires TextShaper.cs to be declared as partial
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    public partial class TextShaper
    {
        #region Vertical Text Shaping (Excel text rotation value 255)

        /// <summary>
        /// Shapes text for vertical layout (top-to-bottom glyph stacking).
        /// Used for Excel vertical text mode (text rotation value 255).
        /// No GSUB or GPOS is applied - the pipeline is intentionally minimal:
        /// character mapping + vertical metrics lookup only.
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <param name="options">Shaping options (reserved for future use)</param>
        /// <returns>Shaped vertical text with positioned glyphs in font design units</returns>
        public ShapedVerticalText ShapeVertical(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedVerticalText
                {
                    OriginalText = text ?? string.Empty,
                    Glyphs = new VerticalShapedGlyph[0]
                };
            }

            var glyphs = MapToVerticalGlyphs(text);

            return new ShapedVerticalText
            {
                OriginalText = text,
                Glyphs = glyphs.ToArray()
            };
        }

        /// <summary>
        /// Shapes text for vertical layout into lightweight VerticalGlyphHeight structs
        /// optimized for vertical text measurement.
        /// Analogous to ShapeLight() for horizontal text.
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <param name="options">Shaping options (reserved for future use)</param>
        /// <returns>Array of lightweight vertical glyph height structs (8 bytes each)</returns>
        public VerticalGlyphHeight[] ShapeLightVertical(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new VerticalGlyphHeight[0];
            }

            var glyphs = MapToVerticalGlyphs(text);
            return ExtractVerticalGlyphHeights(glyphs);
        }

        /// <summary>
        /// Maps characters to vertical glyphs, resolving advance heights from vmtx.
        /// Falls back to hmtx advance widths for fonts without vertical metrics.
        /// Handles surrogate pairs for supplementary plane characters.
        /// </summary>
        private List<VerticalShapedGlyph> MapToVerticalGlyphs(string text)
        {
            var glyphs = new List<VerticalShapedGlyph>(text.Length);

            int i = 0;
            while (i < text.Length)
            {
                uint codePoint;
                int charCount;

                if (i < text.Length - 1 && char.IsHighSurrogate(text[i]) && char.IsLowSurrogate(text[i + 1]))
                {
                    codePoint = (uint)char.ConvertToUtf32(text[i], text[i + 1]);
                    charCount = 2;
                }
                else if (char.IsSurrogate(text[i]))
                {
                    // Lone surrogate - map to .notdef
                    codePoint = 0;
                    charCount = 1;
                }
                else
                {
                    codePoint = text[i];
                    charCount = 1;
                }

                // Resolve glyph via font provider (with fallback support)
                OpenTypeFont font;
                ushort glyphId;
                _fontProvider.TryGetGlyphFont(codePoint, out font, out glyphId);

                byte fontId = GetOrRegisterFontId(font);

                ushort advanceHeight;
                short topSideBearing;
                GetVerticalMetrics(font, glyphId, out advanceHeight, out topSideBearing);
                // Fetch horizontal advance width for centering
                ushort advanceWidth = font.HmtxTable.GetAdvanceWidth(glyphId);

                glyphs.Add(new VerticalShapedGlyph(
                    glyphId,
                    advanceHeight,
                    topSideBearing,
                    advanceWidth,
                    (ushort)i,
                    (byte)charCount,
                    fontId
                ));

                i += charCount;
            }

            return glyphs;
        }

        /// <summary>
        /// Extracts VerticalGlyphHeight structs from a list of VerticalShapedGlyphs.
        /// Discards TopSideBearing as it is not needed for height measurement.
        /// </summary>
        private VerticalGlyphHeight[] ExtractVerticalGlyphHeights(List<VerticalShapedGlyph> glyphs)
        {
            var result = new VerticalGlyphHeight[glyphs.Count];
            for (int i = 0; i < glyphs.Count; i++)
            {
                var g = glyphs[i];
                result[i] = new VerticalGlyphHeight
                {
                    YAdvance = g.YAdvance,
                    ClusterIndex = g.ClusterIndex,
                    CharCount = g.CharCount
                };
            }
            return result;
        }

        /// <summary>
        /// Resolves vertical metrics for a glyph from the given font.
        /// Uses vmtx (advanceHeight + topSideBearing) when available.
        /// Falls back to hmtx advanceWidth as advanceHeight when vmtx is absent,
        /// with topSideBearing set to 0.
        /// </summary>
        private void GetVerticalMetrics(OpenTypeFont font, ushort glyphId,
                                        out ushort advanceHeight, out short topSideBearing)
        {
            var vmtx = font.VmtxTable;
            if (vmtx != null)
            {
                advanceHeight = vmtx.GetAdvanceHeight(glyphId);
                topSideBearing = vmtx.GetTopSideBearing(glyphId);
            }
            else
            {
                // Fallback: use horizontal advance width as advance height.
                // Gives reasonable glyph spacing for fonts without vertical metrics.
                advanceHeight = font.HmtxTable.GetAdvanceWidth(glyphId);
                topSideBearing = 0;
            }
        }

        #endregion
    }
}