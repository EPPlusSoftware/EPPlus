/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage
{
    public abstract class CoverageTable : FontTableElement
    {
        public ushort CoverageFormat { get; set; }
        public abstract ushort[] CoveredGlyphs { get; }

        public abstract int GetGlyphIndex(ushort glyphId);

        public abstract ushort[] GetCoveredGlyphs();

        /// <summary>
        /// Checks if a glyph ID is covered by this coverage table.
        /// </summary>
        /// <param name="glyphId">The glyph ID to check</param>
        /// <returns>True if the glyph is covered, false otherwise</returns>
        public bool IsCovered(ushort glyphId)
        {
            return GetGlyphIndex(glyphId) >= 0;
        }
    }
}
