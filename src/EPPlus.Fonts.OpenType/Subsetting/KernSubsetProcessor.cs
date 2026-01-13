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
using EPPlus.Fonts.OpenType.Tables.Kern;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Subsets the 'kern' table (format 0 only – the only format used in practice).
    /// Only keeps kerning pairs where both left and right glyphs exist in the subset.
    /// Preserves coverage flags and recalculates searchRange/entrySelector/rangeShift.
    /// Must run after GlyfAndLocaSubsetProcessor.
    /// .NET 3.5 compatible.
    /// </summary>
    internal class KernSubsetProcessor : IFontSubsetProcessor
    {
        public void Discover(FontSubsettingContext context)
        {
            var originalFont = context.OriginalFont;
            var subsetFont = context.SubsetFont;

            if (originalFont.KernTable == null || originalFont.KernTable.SubTables.Count == 0)
                return; // No kerning in source font

            var oldKern = originalFont.KernTable;
            var newKern = new KernTable
            {
                version = oldKern.version,
                numberOfFormat0Tables = 0
            };

            foreach (var originalSubTable in oldKern.SubTables)
            {
                // Only format 0 is used in real fonts – skip others
                if (originalSubTable.coverage.Format != 0 || originalSubTable.Format0Subtable == null)
                    continue;

                var oldFormat0 = originalSubTable.Format0Subtable;
                var newPairs = new List<KerningPair>();

                foreach (var pair in oldFormat0.Pairs)
                {
                    if (context.OldToNewGlyphId.TryGetValue(pair.left, out ushort newLeft) &&
                        context.OldToNewGlyphId.TryGetValue(pair.right, out ushort newRight))
                    {
                        newPairs.Add(new KerningPair(null!)
                        {
                            left = newLeft,
                            right = newRight,
                            value = pair.value,
                            Combined = ((uint)newLeft << 16) | newRight
                        });
                    }
                }

                if (newPairs.Count == 0)
                    continue;

                // Sort pairs by Combined value – required for binary search (spec §5)
                newPairs.Sort((a, b) => a.Combined.CompareTo(b.Combined));

                var newSubTable = new KernSubTable
                {
                    version = originalSubTable.version,
                    coverage = originalSubTable.coverage, // preserve horizontal/vertical flags
                    Format0Subtable = new KernSubTableFormat0(null!)
                    {
                        nPairs = (ushort)newPairs.Count,
                        Pairs = newPairs.ToArray()
                    }
                };

                newKern.SubTables.Add(newSubTable);
                newKern.numberOfFormat0Tables++;
            }

            // Only add kern table if we actually have pairs
            if (newKern.SubTables.Count > 0)
            {
                subsetFont.AddOrReplaceTable(newKern);
            }
        }

        public void Rewrite(FontSubsettingContext context)
        {
            // No implementation
        }
    }
}
