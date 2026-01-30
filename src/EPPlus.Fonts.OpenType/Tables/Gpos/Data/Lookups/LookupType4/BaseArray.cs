/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS BaseArray (Type 4)
 *************************************************************************************************/

using EPPlus.Fonts.OpenType.Tables.Name;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4
{
    /// <summary>
    /// Array of BaseRecords defining attachment points for base glyphs (letters).
    /// One record per base glyph in coverage order.
    /// </summary>
    public class BaseArray
    {
        /// <summary>
        /// Number of BaseRecords in the array
        /// </summary>
        public ushort BaseCount { get; internal set; }

        /// <summary>
        /// Array of BaseRecords, one per base glyph in coverage order
        /// </summary>
        public BaseRecord[] Records { get; internal set; }
    }
}