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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data;
using System.Linq;

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
