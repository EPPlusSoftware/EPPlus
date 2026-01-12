/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS BaseRecord (Type 4)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4
{
    /// <summary>
    /// Record defining attachment points for a base glyph (letter).
    /// Contains one anchor per mark class.
    /// </summary>
    public class BaseRecord
    {
        /// <summary>
        /// Array of offsets to Anchor tables (from beginning of BaseArray)
        /// One offset per mark class (length = MarkClassCount from subtable)
        /// </summary>
        public ushort[] BaseAnchorOffsets { get; internal set; }

        /// <summary>
        /// Array of Anchor tables defining attachment points for each mark class.
        /// Index corresponds to mark class (e.g., [0] = top accents, [1] = bottom accents)
        /// </summary>
        public AnchorTable[] BaseAnchors { get; internal set; }
    }
}