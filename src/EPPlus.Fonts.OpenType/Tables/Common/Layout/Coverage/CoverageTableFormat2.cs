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
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage.IO;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage
{
    /// <summary>
    /// Represents a Coverage Table Format 2, which defines glyph coverage using ranges.
    /// </summary>
    public class CoverageTableFormat2 : CoverageTable
    {
        /// <summary>
        /// Gets or sets the number of range records.
        /// </summary>
        public ushort RangeCount { get; set; }

        /// <summary>
        /// Gets or sets the list of range records defining the covered glyphs.
        /// </summary>
        public List<CoverageRangeRecord> RangeRecords { get; set; } = new List<CoverageRangeRecord>();

        /// <summary>
        /// Gets an array of all Glyph IDs covered by this table.
        /// </summary>
        public override ushort[] CoveredGlyphs
        {
            get
            {
                return GetCoveredGlyphs();
            }
        }

        /// <summary>
        /// Returns the coverage index for a specific Glyph ID.
        /// </summary>
        /// <param name="glyphId">The Glyph ID to look up.</param>
        /// <returns>The coverage index, or -1 if the glyph is not covered.</returns>
        public override int GetGlyphIndex(ushort glyphId)
        {
            if (RangeRecords == null || RangeRecords.Count == 0) return -1;

            // Binary search through the ranges for performance (O(log n))
            int low = 0;
            int high = RangeRecords.Count - 1;

            while (low <= high)
            {
                int mid = low + (high - low) / 2;
                var range = RangeRecords[mid];

                if (glyphId >= range.StartGlyphID && glyphId <= range.EndGlyphID)
                {
                    // The Coverage Index is calculated as: StartCoverageIndex + (GlyphID - StartGlyphID)
                    return range.StartCoverageIndex + (glyphId - range.StartGlyphID);
                }

                if (glyphId < range.StartGlyphID)
                {
                    high = mid - 1;
                }
                else
                {
                    low = mid + 1;
                }
            }

            return -1; // Glyph ID not found in any range
        }

        /// <summary>
        /// Generates an array of all covered Glyph IDs by flattening the range records.
        /// </summary>
        public override ushort[] GetCoveredGlyphs()
        {
            if (RangeRecords == null) return new ushort[0];

            // Uses SelectMany to flatten the ranges produced by GetRange
            return RangeRecords.SelectMany(r => GetRange(r.StartGlyphID, r.EndGlyphID)).ToArray();
        }

        /// <summary>
        /// Helper to generate a sequence of Glyph IDs between start and end inclusive.
        /// Compatible with .NET 3.5.
        /// </summary>
        private IEnumerable<ushort> GetRange(ushort start, ushort end)
        {
            if (start > end)
            {
                yield break;
            }

            for (int i = start; i <= end; i++)
            {
                yield return (ushort)i;
            }
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            new CoverageTableFormat2Serializer().Serialize(this, writer);
        }

        /// <summary>
        /// Factory method to create a Format 2 Coverage Table from a list of sorted Glyph IDs.
        /// This method automatically groups consecutive IDs into ranges to optimize size.
        /// </summary>
        /// <param name="newGlyphs">A sorted list of Glyph IDs.</param>
        internal static CoverageTableFormat2 CreateCoverageFormat2(List<ushort> newGlyphs)
        {
            CoverageTableFormat2 coverage = new CoverageTableFormat2();
            if (newGlyphs == null || newGlyphs.Count == 0)
            {
                coverage.RangeCount = 0;
                return coverage;
            }

            ushort startGlyph = newGlyphs[0];
            ushort lastGlyph = newGlyphs[0];
            ushort startCoverageIndex = 0;

            for (int i = 1; i <= newGlyphs.Count; i++)
            {
                // Close the current range if:
                // 1. We reached the end of the list
                // 2. The current GID is not consecutive (there is a gap)
                if (i == newGlyphs.Count || newGlyphs[i] != lastGlyph + 1)
                {
                    coverage.RangeRecords.Add(new CoverageRangeRecord
                    {
                        StartGlyphID = startGlyph,
                        EndGlyphID = lastGlyph,
                        StartCoverageIndex = startCoverageIndex
                    });

                    if (i < newGlyphs.Count)
                    {
                        startGlyph = newGlyphs[i];
                        lastGlyph = newGlyphs[i];
                        // The next coverage index is the number of glyphs processed so far
                        startCoverageIndex = (ushort)i;
                    }
                }
                else
                {
                    lastGlyph = newGlyphs[i];
                }
            }

            coverage.RangeCount = (ushort)coverage.RangeRecords.Count;
            return coverage;
        }
    }
}