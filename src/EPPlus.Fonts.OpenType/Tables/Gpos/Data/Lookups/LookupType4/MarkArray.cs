/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS MarkArray (Type 4, 5, 6)
 *************************************************************************************************/

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4
{
    /// <summary>
    /// Array of MarkRecords defining attachment points for mark glyphs (accents).
    /// Used by MarkToBase, MarkToLigature, and MarkToMark lookups.
    /// </summary>
    public class MarkArray
    {
        /// <summary>
        /// Number of MarkRecords in the array
        /// </summary>
        public ushort MarkCount { get; internal set; }

        /// <summary>
        /// Array of MarkRecords, one per mark glyph in coverage order
        /// </summary>
        public MarkRecord[] Records { get; internal set; }
    }
}