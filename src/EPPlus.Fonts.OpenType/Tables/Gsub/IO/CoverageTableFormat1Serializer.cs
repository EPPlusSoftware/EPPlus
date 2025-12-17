using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    internal class CoverageTableFormat1Serializer
    {
        public void Serialize(CoverageTableFormat1 table, FontsBinaryWriter writer)
        {
            // The table should only be serialized if it contains data
            if (table.GlyphArray == null || table.GlyphArray.Length == 0)
            {
                // The caller (LigatureSubstSubTableSerializer) must handle the offset being 0.
                return;
            }

            // Must ensure the array is sorted ascendingly as required by the spec
            // Since our CreateSubset already sorts by the new IDs, this should be safe.
            ushort[] sortedGlyphArray = table.GlyphArray.OrderBy(g => g).ToArray();

            // USHORT CoverageFormat (1)
            writer.WriteUInt16BigEndian(1);

            // USHORT GlyphCount
            writer.WriteUInt16BigEndian((ushort)sortedGlyphArray.Length);

            // USHORT[] GlyphArray
            foreach (ushort gid in sortedGlyphArray)
            {
                writer.WriteUInt16BigEndian(gid);
            }
        }
    }
}
