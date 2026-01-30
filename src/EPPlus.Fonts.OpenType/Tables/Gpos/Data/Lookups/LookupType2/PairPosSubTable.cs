/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           GPOS PairPos subtable (Type 2)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2
{
    /// <summary>
    /// GPOS Lookup Type 2: Pair Adjustment Positioning Subtable.
    /// Used for kerning - adjusting spacing between specific glyph pairs.
    /// Abstract base for Format 1 and Format 2.
    /// </summary>
    public abstract class PairPosSubTable : GposSubTableBase
    {
        /// <summary>
        /// Subtable format (1 or 2)
        /// </summary>
        public ushort SubtableFormat { get; set; }

        /// <summary>
        /// Coverage table - which glyphs are affected
        /// </summary>
        public CoverageTable Coverage { get; set; }

        /// <summary>
        /// Value format for first glyph in pair
        /// </summary>
        public ushort ValueFormat1 { get; set; }

        /// <summary>
        /// Value format for second glyph in pair
        /// </summary>
        public ushort ValueFormat2 { get; set; }


        /// <summary>
        /// Gets the positioning adjustment for a specific glyph pair
        /// </summary>
        /// <param name="firstGlyph">First glyph in pair</param>
        /// <param name="secondGlyph">Second glyph in pair</param>
        /// <param name="value1">Adjustment for first glyph (output)</param>
        /// <param name="value2">Adjustment for second glyph (output)</param>
        /// <returns>True if pair has positioning data</returns>
        public abstract bool TryGetPairAdjustment(
            ushort firstGlyph,
            ushort secondGlyph,
            out ValueRecord value1,
            out ValueRecord value2);
    }
}