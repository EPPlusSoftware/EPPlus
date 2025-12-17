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

using EPPlus.Fonts.OpenType.Tables.Hmtx;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting.Processors
{
    /// <summary>
    /// Creates the final hmtx table for the subset font.
    /// Uses the glyph ID remapping from GlyfAndLocaSubsetProcessor.
    /// Must run after GlyfAndLocaSubsetProcessor.
    /// .NET 3.5 compatible.
    /// </summary>
    internal class HmtxSubsetProcessor : IFontSubsetProcessor
    {
        public void Discover(FontSubsettingContext context)
        {
            // No implementation
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var originalHmtx = context.OriginalFont.HmtxTable;
            if (originalHmtx == null) return;

            int finalGlyphCount = context.NewToOldGlyphId.Count;

            // Create the new metrics storage
            var newHMetrics = new List<LongHorMetric>(finalGlyphCount);

            // In our subset, we simplify by making numberOfHMetrics equal to numGlyphs
            // This means we only use the hMetrics list, and leftSideBearings will be empty.
            for (int i = 0; i < finalGlyphCount; i++)
            {
                ushort oldGlyphId = context.NewToOldGlyphId[i];

                newHMetrics.Add(new LongHorMetric
                {
                    advanceWidth = originalHmtx.GetAdvanceWidth(oldGlyphId),
                    lsb = originalHmtx.GetLeftSideBearing(oldGlyphId)
                });
            }

            var newHmtx = new HmtxTable(newHMetrics);
            context.SubsetFont.AddOrReplaceTable(newHmtx);
        }
    }
}

