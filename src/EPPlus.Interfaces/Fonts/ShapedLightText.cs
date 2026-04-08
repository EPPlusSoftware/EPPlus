/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/18/2026         EPPlus Software AB           Lightweight multi-font aware shaping result
 *************************************************************************************************/
using System.Diagnostics;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Lightweight result of text shaping, optimized for measurement and wrapping.
    /// Contains <see cref="GlyphWidth"/> structs (8 bytes each) instead of full
    /// <see cref="ShapedGlyph"/> objects, plus per-font metrics inherited from
    /// <see cref="ShapedTextBase"/> for correct multi-font calculation.
    /// </summary>
    [DebuggerDisplay("Glyphs: {Glyphs.Length}, Fonts: {FontUnitsPerEm?.Length ?? 0}")]
    public class ShapedLightText : ShapedTextBase
    {
        /// <summary>
        /// Array of lightweight glyph widths with positioning information.
        /// </summary>
        public GlyphWidth[] Glyphs { get; set; }

        protected override int GetGlyphCount() =>
            Glyphs?.Length ?? 0;

        protected override int GetGlyphXAdvance(int index) =>
            Glyphs[index].XAdvance;

        protected override byte GetGlyphFontId(int index) =>
            Glyphs[index].FontId;

        /// <summary>
        /// Fills a character width buffer with per-character widths in points.
        /// Handles multi-font glyphs correctly.
        /// </summary>
        public void FillCharWidths(float fontSize, double[] charWidths, int textLength)
        {
            if (Glyphs == null) return;

            bool singleFont = FontUnitsPerEm != null && FontUnitsPerEm.Length == 1;
            float singleScale = singleFont ? fontSize / FontUnitsPerEm[0] : 0f;

            foreach (var glyph in Glyphs)
            {
                int idx = glyph.ClusterIndex;
                if (idx >= 0 && idx < textLength)
                {
                    if (singleFont)
                    {
                        charWidths[idx] += glyph.XAdvance * singleScale;
                    }
                    else
                    {
                        float upm = (FontUnitsPerEm != null && glyph.FontId < FontUnitsPerEm.Length)
                            ? FontUnitsPerEm[glyph.FontId]
                            : 1000f;
                        charWidths[idx] += glyph.XAdvance * (fontSize / upm);
                    }
                }
            }
        }
    }
}