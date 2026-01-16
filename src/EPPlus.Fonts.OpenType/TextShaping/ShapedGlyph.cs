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

namespace EPPlus.Fonts.OpenType.TextShaping
{
    /// <summary>
    /// Represents a shaped glyph with positioning information.
    /// All measurements are in font units (not PDF points or pixels).
    /// </summary>
    public class ShapedGlyph
    {
        public ShapedGlyph()
        {
            
        }
        public ShapedGlyph(ushort glyphId, int xAdvance)
        {
            GlyphId = glyphId;
            XAdvance = xAdvance;
            YAdvance = 0;
            XOffset = 0;
            YOffset = 0;
            ClusterIndex = 0;
            CharCount = 1;
        }

        /// <summary>
        /// The glyph ID in the font.
        /// </summary>
        public ushort GlyphId { get; set; }

        /// <summary>
        /// Horizontal advance width in font units.
        /// This includes any kerning adjustments from GPOS.
        /// </summary>
        public int XAdvance { get; set; }

        /// <summary>
        /// Vertical advance height in font units.
        /// Typically 0 for horizontal text.
        /// </summary>
        public int YAdvance { get; set; }

        /// <summary>
        /// Horizontal offset adjustment in font units.
        /// Used for positioning marks, subscripts, superscripts.
        /// </summary>
        public int XOffset { get; set; }

        /// <summary>
        /// Vertical offset adjustment in font units.
        /// Used for positioning marks, subscripts, superscripts.
        /// </summary>
        public int YOffset { get; set; }

        /// <summary>
        /// Index of the original character(s) that produced this glyph.
        /// Used for text selection and editing.
        /// For ligatures, this points to the first character.
        /// </summary>
        public int ClusterIndex { get; set; }

        /// <summary>
        /// Number of characters consumed by this glyph.
        /// 1 for normal glyphs, 2+ for ligatures (e.g., "fi" → 1 glyph, 2 chars).
        /// </summary>
        public int CharCount { get; set; }
    }
}