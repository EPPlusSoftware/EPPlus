using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class CoverageTableFormat2 : CoverageTable
    {
        public ushort RangeCount { get; set; }
        public List<CoverageRangeRecord> RangeRecords { get; set; } = new List<CoverageRangeRecord>();

        public override ushort[] CoveredGlyphs
        {
            get
            {
                // Använd SelectMany för att platta ut resultatet från vår GetRange metod
                // Denna syntax (SelectMany) är tillgänglig i .NET 3.5.
                return RangeRecords.SelectMany(r => GetRange(r.StartGlyphID, r.EndGlyphID)).ToArray();
            }
        }

        /// <summary>
        /// .NET 3.5 compatible method to generate an integer sequence (equivalent to Enumerable.Range).
        /// Since Start and End Glyph IDs are USHORT, the range will not exceed 65535 elements, 
        /// which fits within a standard int iterator.
        /// </summary>
        /// <param name="start">The starting Glyph ID.</param>
        /// <param name="end">The ending Glyph ID.</param>
        /// <returns>An IEnumerable<ushort> representing the range [start, end].</returns>
        private IEnumerable<ushort> GetRange(ushort start, ushort end)
        {
            // If start > end, we should ideally log an error or return empty, but for robust reading:
            if (start > end)
            {
                yield break;
            }

            // Loop from start Glyph ID up to and including the end Glyph ID
            for (int i = start; i <= end; i++)
            {
                // Yield return the current value as USHORT
                yield return (ushort)i;
            }
        }
    }
}
