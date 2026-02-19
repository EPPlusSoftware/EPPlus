
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           vmtx table implementation (vertical text support)
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Tables.Vmtx
{
    /// <summary>
    /// Represents a single entry in the vmtx table's longVerMetric array.
    /// Analogous to LongHorMetric in hmtx.
    /// </summary>
    public class LongVerMetric
    {
        /// <summary>
        /// The advance height of the glyph in font design units.
        /// </summary>
        public ushort AdvanceHeight { get; set; }

        /// <summary>
        /// The top side bearing of the glyph in font design units.
        /// Distance from the vertical origin to the top of the glyph bounding box.
        /// </summary>
        public short TopSideBearing { get; set; }
    }
}
