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
using EPPlus.Fonts.OpenType.Tables.Gsub.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    public class CoverageTableFormat1 : CoverageTable
    {
        public ushort GlyphCount { get; set; }
        public ushort[] GlyphArray { get; set; }
        public override ushort[] CoveredGlyphs => GlyphArray;

        public override ushort[] GetCoveredGlyphs()
        {
            // Format 1 is already a flat array
            return GlyphArray ?? new ushort[0];
        }

        public override int GetGlyphIndex(ushort glyphId)
        {
            if (GlyphArray == null || GlyphArray.Length == 0) return -1;

            // Binary search is efficient as GlyphArray MUST be sorted ascending
            int low = 0;
            int high = GlyphArray.Length - 1;

            while (low <= high)
            {
                int mid = low + (high - low) / 2;
                ushort midVal = GlyphArray[mid];

                if (midVal == glyphId) return mid;
                if (midVal < glyphId) low = mid + 1;
                else high = mid - 1;
            }

            return -1; // Not found
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            new CoverageTableFormat1Serializer().Serialize(this, writer);
        }
    }
}
