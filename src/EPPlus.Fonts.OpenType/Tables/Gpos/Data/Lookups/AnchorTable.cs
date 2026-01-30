/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Anchor table (shared across lookup types)
 *************************************************************************************************/

using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups
{
    /// <summary>
    /// Anchor table defining a position attachment point.
    /// Used by MarkToBase, MarkToLigature, MarkToMark, and Cursive attachment.
    /// Three formats exist, but we implement the most common ones.
    /// </summary>
    public class AnchorTable
    {
        /// <summary>
        /// Format identifier (1, 2, or 3)
        /// </summary>
        public ushort AnchorFormat { get; internal set; }

        /// <summary>
        /// X coordinate of anchor point (in design units)
        /// </summary>
        public short XCoordinate { get; internal set; }

        /// <summary>
        /// Y coordinate of anchor point (in design units)
        /// </summary>
        public short YCoordinate { get; internal set; }

        /// <summary>
        /// Anchor point index (Format 2 only)
        /// Used to identify which contour point to use
        /// </summary>
        public ushort AnchorPoint { get; internal set; }

        /// <summary>
        /// Offset to Device table for X coordinate (Format 3 only)
        /// We don't implement device tables yet - stored for completeness
        /// </summary>
        public ushort XDeviceOffset { get; internal set; }

        /// <summary>
        /// Offset to Device table for Y coordinate (Format 3 only)
        /// We don't implement device tables yet - stored for completeness
        /// </summary>
        public ushort YDeviceOffset { get; internal set; }
    }
}