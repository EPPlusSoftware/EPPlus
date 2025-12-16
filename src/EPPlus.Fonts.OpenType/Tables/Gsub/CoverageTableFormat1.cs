using EPPlus.Fonts.OpenType.Tables.Gsub.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
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
