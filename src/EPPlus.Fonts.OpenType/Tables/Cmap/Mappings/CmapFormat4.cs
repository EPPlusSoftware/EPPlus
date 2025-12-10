using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Mappings
{
    internal static class CmapFormat4
    {
        /// <summary>
        /// Creates a minimal and correct Format 4 cmap subtable from a dictionary of Unicode code points to glyph IDs.
        /// Optimized for font subsetting – only includes used characters and guarantees .notdef (code point 0 to glyph 0).
        /// Fully compatible with .NET 3.5 and EPPlus coding standards.
        /// </summary>
        /// <param name="codePointToGlyphId">Mapping of Unicode code points to new (subset) glyph IDs</param>
        /// <returns>A fully populated CmapSubtable4 ready for serialization</returns>
        internal static CmapSubtable4 CreateFromMappings(Dictionary<uint, ushort> codePointToGlyphId)
        {
            if (codePointToGlyphId == null || codePointToGlyphId.Count == 0)
                throw new ArgumentException("No character-to-glyph mappings provided.", "codePointToGlyphId");

            // Ensure .notdef mapping exists (U+0000 to glyph 0)
            codePointToGlyphId[0] = 0;

            // Sort code points for sequential processing
            List<uint> sortedCodes = new List<uint>(codePointToGlyphId.Keys);
            sortedCodes.Sort();

            // Temporary lists to build segments
            List<ushort> segStart = new List<ushort>();
            List<ushort> segEnd = new List<ushort>();
            List<short> segDelta = new List<short>();
            List<ushort> segRangeOffset = new List<ushort>();
            List<ushort> glyphIdArray = new List<ushort>();

            uint currentStart = sortedCodes[0];
            uint currentEnd = sortedCodes[0];
            ushort currentFirstGlyph = codePointToGlyphId[currentStart];

            for (int i = 1; i < sortedCodes.Count; i++)
            {
                uint code = sortedCodes[i];
                ushort glyph = codePointToGlyphId[code];

                if (code == currentEnd + 1 && glyph == (ushort)(currentFirstGlyph + (code - currentStart)))
                {
                    // Continue current segment
                    currentEnd = code;
                }
                else
                {
                    // Finalize current segment
                    AddSegment(segStart, segEnd, segDelta, segRangeOffset, glyphIdArray,
                               currentStart, currentEnd, currentFirstGlyph, codePointToGlyphId);

                    // Start new segment
                    currentStart = code;
                    currentEnd = code;
                    currentFirstGlyph = glyph;
                }
            }

            // Add final segment
            AddSegment(segStart, segEnd, segDelta, segRangeOffset, glyphIdArray,
                       currentStart, currentEnd, currentFirstGlyph, codePointToGlyphId);

            // Add terminating segment (0xFFFF)
            segStart.Add(0xFFFF);
            segEnd.Add(0xFFFF);
            segDelta.Add(1);           // 0xFFFF + 1 to glyph 0
            segRangeOffset.Add(0);

            int segCount = segStart.Count;

            // Calculate search parameters
            int pow2 = 1;
            while (pow2 * 2 <= segCount) pow2 *= 2;
            ushort searchRange = (ushort)(pow2 * 2);
            ushort entrySelector = 0;
            int temp = pow2;
            while (temp > 1)
            {
                temp /= 2;
                entrySelector++;
            }
            ushort rangeShift = (ushort)(segCount * 2 - searchRange);
            
            // Calculate correct table length (header + data + 4-byte alignment)
            int headerSize = 16; // format(2) + length(2) + language(2) + segCountX2(2) + searchRange(2) + entrySelector(2) + rangeShift(2)
            int segmentDataSize = segCount * 8; // endCode, startCode, idDelta, idRangeOffset – 2 bytes each × segCount
            int glyphIdArraySize = glyphIdArray.Count * 2;
            int totalSize = headerSize + segmentDataSize + glyphIdArraySize;

            // 4-byte alignment required by OpenType spec
            int paddedSize = (totalSize + 3) & ~3;

            return new CmapSubtable4
            {
                SegCountX2 = (ushort)(segCount * 2),
                SearchRange = searchRange,
                EntrySelector = entrySelector,
                RangeShift = rangeShift,
                EndCode = segEnd.ToArray(),
                ReservedPad = 0,
                StartCode = segStart.ToArray(),
                IdDelta = segDelta.ToArray(),
                IdRangeOffset = segRangeOffset.ToArray(),
                GlyphIdArray = glyphIdArray.ToArray(),
                Length = (uint)paddedSize
            };
        }

        // Helper: adds one segment to the format 4 arrays
        // I klassen: internal static class CmapFormat4
        // Ersätt hela den befintliga AddSegment-metoden med denna:

        private static void AddSegment(
            List<ushort> startCodes,
            List<ushort> endCodes,
            List<short> deltas,
            List<ushort> rangeOffsets,
            List<ushort> glyphIdArray,
            uint startCode,
            uint endCode,
            ushort firstGlyphId,
            Dictionary<uint, ushort> mapping)
        {
            startCodes.Add((ushort)startCode);
            endCodes.Add((ushort)endCode);

            // Försök använda idDelta (sekventiell mappning)
            bool isSequential = true;
            for (uint c = startCode; c <= endCode; c++)
            {
                ushort expectedGlyph = (ushort)((firstGlyphId + (c - startCode)) & 0xFFFF);
                if (mapping[c] != expectedGlyph)
                {
                    isSequential = false;
                    break;
                }
            }

            if (isSequential)
            {
                // Sekventiell → använd bara idDelta
                short delta = (short)((int)firstGlyphId - (int)startCode);
                deltas.Add(delta);
                rangeOffsets.Add(0);
            }
            else
            {
                // Icke-sekventiell → använd glyphIdArray
                int glyphArrayStartIndex = glyphIdArray.Count;

                // Lägg till alla glyph-ID:n för detta segment
                for (uint c = startCode; c <= endCode; c++)
                {
                    glyphIdArray.Add(mapping[c]);
                }

                // RÄTT BERÄKNING AV idRangeOffset (detta var felet!)
                // Avstånd i bytes från denna idRangeOffset-post till glyphIdArray[startCode]
                int entriesFromThisIncludingThis = rangeOffsets.Count + 1;
                int idRangeOffset = entriesFromThisIncludingThis * 2 + (glyphIdArray.Count - glyphArrayStartIndex) * 2 - 2;

                deltas.Add(0);                    // idDelta ignoreras när idRangeOffset != 0
                rangeOffsets.Add((ushort)idRangeOffset);
            }
        }
    }
}
