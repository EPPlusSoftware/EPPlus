/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/24/2026         EPPlus Software AB           Lightweight glyph for text measurement
 *************************************************************************************************/
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Lightweight glyph representation optimized for text measurement and wrapping.
    /// Contains only essential data needed for width calculations.
    /// This struct is 8 bytes total - 85% smaller than ShapedGlyph class (56 bytes).
    /// </summary>
    [DebuggerDisplay("XAdvance: {XAdvance}, ClusterIndex: {ClusterIndex}, CharCount: {CharCount}")]
    [StructLayout(LayoutKind.Sequential)]
    public struct GlyphWidth
    {
        /// <summary>
        /// Horizontal advance width in font units.
        /// Includes kerning adjustments.
        /// Max value: 65,535 font units (more than sufficient for any font).
        /// </summary>
        public ushort XAdvance;

        /// <summary>
        /// Index of the original character(s) that produced this glyph.
        /// Used to map glyph widths back to character positions.
        /// For ligatures, this points to the first character.
        /// Max value: 65,535 characters per string.
        /// </summary>
        public ushort ClusterIndex;

        /// <summary>
        /// Number of characters consumed by this glyph.
        /// 1 for normal glyphs, 2+ for ligatures (e.g., "fi" → 1 glyph, 2 chars).
        /// Max value: 255 characters per ligature (more than sufficient).
        /// </summary>
        public byte CharCount;

        /// <summary>
        /// Which font produced this glyph (0 = first used font, 1+ = fallbacks).
        /// Needed for correct point conversion when fonts have different UnitsPerEm.
        /// </summary>
        public byte FontId;

    }
}