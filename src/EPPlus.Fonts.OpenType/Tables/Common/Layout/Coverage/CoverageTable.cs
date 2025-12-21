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
namespace EPPlus.Fonts.OpenType.Tables.Common.Coverage
{
    public abstract class CoverageTable : FontTableElement
    {
        public ushort CoverageFormat { get; set; }
        public abstract ushort[] CoveredGlyphs { get; }

        public abstract int GetGlyphIndex(ushort glyphId);

        public abstract ushort[] GetCoveredGlyphs();
    }
}
