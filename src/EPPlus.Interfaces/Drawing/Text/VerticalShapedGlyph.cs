/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           Vertical text shaping support
 *************************************************************************************************/
using System.Diagnostics;

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// Represents a shaped glyph with vertical positioning information.
    /// Used for Excel vertical text mode (text rotation value 255).
    /// All measurements are in font design units (not PDF points or pixels).
    /// Analogous to <see cref="ShapedGlyph"/> for horizontal text.
    /// </summary>
    [DebuggerDisplay("GlyphId: {GlyphId}, YAdvance: {YAdvance}, CharCount: {CharCount}")]
    public class VerticalShapedGlyph
    {
        /// <summary>
        /// The glyph ID in the font.
        /// Range: 0-65,535.
        /// </summary>
        public ushort GlyphId;

        /// <summary>
        /// Vertical advance height in font design units.
        /// Sourced from the 'vmtx' table (advanceHeight).
        /// Falls back to advanceWidth from 'hmtx' if 'vmtx' is not present in the font.
        /// </summary>
        public ushort YAdvance;

        /// <summary>
        /// Top side bearing in font design units.
        /// Distance from the vertical origin to the top of the glyph bounding box.
        /// Sourced from the 'vmtx' table (topSideBearing).
        /// </summary>
        public short TopSideBearing;

        /// <summary>
        /// Index of the original character that produced this glyph.
        /// Used to map shaped glyphs back to character positions.
        /// Range: 0-65,535 characters per string.
        /// </summary>
        public ushort ClusterIndex;

        /// <summary>
        /// Number of characters consumed by this glyph.
        /// Typically 1 for vertical text (no ligatures in vertical pipeline).
        /// </summary>
        public byte CharCount;

        /// <summary>
        /// ID of the font that provided this glyph.
        /// 0 = primary font, 1+ = fallback fonts.
        /// </summary>
        public byte FontId;

        // Total size: 10 bytes + 2 bytes padding = 12 bytes

        /// <summary>
        /// Creates a new vertical shaped glyph with default values.
        /// </summary>
        public VerticalShapedGlyph()
        {
            CharCount = 1;
        }

        /// <summary>
        /// Creates a new vertical shaped glyph with all fields specified.
        /// </summary>
        public VerticalShapedGlyph(ushort glyphId, ushort yAdvance, short topSideBearing,
                                   ushort clusterIndex, byte charCount, byte fontId = 0)
        {
            GlyphId = glyphId;
            YAdvance = yAdvance;
            TopSideBearing = topSideBearing;
            ClusterIndex = clusterIndex;
            CharCount = charCount;
            FontId = fontId;
        }
    }
}