/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           vmtx subset processor (vertical text support)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Vmtx;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting.Processors
{
    /// <summary>
    /// Creates the subsetted 'vmtx' (Vertical Metrics) table.
    /// Only runs if the original font contains a vmtx table.
    /// Analogous to <see cref="HmtxSubsetProcessor"/> for horizontal metrics.
    /// Must run after GlyfAndLocaSubsetProcessor so that NewToOldGlyphId is populated.
    /// </summary>
    internal class VmtxSubsetProcessor : IFontSubsetProcessor
    {
        public void Discover(FontSubsettingContext context)
        {
            // No additional glyphs to discover - vmtx only carries metrics
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var originalVmtx = context.OriginalFont.VmtxTable;

            // vmtx is optional - skip silently if not present
            if (originalVmtx == null) return;

            int finalGlyphCount = context.NewToOldGlyphId.Count;

            // Simplify: set numberOfVMetrics == numGlyphs (same approach as HmtxSubsetProcessor).
            // This means all entries go into VMetrics and TopSideBearings stays empty.
            var newVMetrics = new List<LongVerMetric>(finalGlyphCount);

            for (int i = 0; i < finalGlyphCount; i++)
            {
                ushort oldGlyphId = context.NewToOldGlyphId[i];

                newVMetrics.Add(new LongVerMetric
                {
                    AdvanceHeight = originalVmtx.GetAdvanceHeight(oldGlyphId),
                    TopSideBearing = originalVmtx.GetTopSideBearing(oldGlyphId)
                });
            }

            var newVmtx = new VmtxTable(newVMetrics);
            context.SubsetFont.AddOrReplaceTable(newVmtx);
        }
    }
}