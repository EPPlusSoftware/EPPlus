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
using System.Runtime.InteropServices;

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// Lightweight glyph representation optimized for vertical text measurement.
    /// Contains only essential data needed for height calculations.
    /// Analogous to <see cref="GlyphWidth"/> for horizontal text.
    /// This struct is 8 bytes total.
    /// </summary>
    [DebuggerDisplay("YAdvance: {YAdvance}, ClusterIndex: {ClusterIndex}, CharCount: {CharCount}")]
    [StructLayout(LayoutKind.Sequential)]
    public struct VerticalGlyphHeight
    {
        /// <summary>
        /// Vertical advance height in font design units.
        /// Sourced from the 'vmtx' table, or falls back to advanceWidth from 'hmtx'.
        /// </summary>
        public ushort YAdvance;

        /// <summary>
        /// Index of the original character that produced this glyph.
        /// Used to map glyph heights back to character positions.
        /// Range: 0-65,535 characters per string.
        /// </summary>
        public ushort ClusterIndex;

        /// <summary>
        /// Number of characters consumed by this glyph.
        /// Typically 1 for vertical text (no ligatures in vertical pipeline).
        /// </summary>
        public byte CharCount;

        // 3 bytes padding to align to 8 bytes total
    }
}