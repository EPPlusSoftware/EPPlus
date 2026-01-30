/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS MarkRecord (Type 4, 5, 6)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4
{
    /// <summary>
    /// Record defining a mark glyph's attachment point and class.
    /// </summary>
    public class MarkRecord
    {
        /// <summary>
        /// Mark class value (which attachment class this mark belongs to)
        /// Example: class 0 = top accents, class 1 = bottom accents
        /// </summary>
        public ushort MarkClass { get; internal set; }

        /// <summary>
        /// Offset to Anchor table for this mark (from beginning of MarkArray)
        /// </summary>
        public ushort MarkAnchorOffset { get; internal set; }

        /// <summary>
        /// Anchor table defining the attachment point on the mark glyph
        /// </summary>
        public AnchorTable MarkAnchor { get; internal set; }
    }
}