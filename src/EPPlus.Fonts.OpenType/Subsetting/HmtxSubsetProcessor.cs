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
        public void Process(FontSubsettingContext context)
        {
            if (context.OriginalFont.HmtxTable == null)
                return; // inget att göra

            var originalHmtx = context.OriginalFont.HmtxTable;
            var newFont = context.SubsetFont;

            int finalGlyphCount = context.NewToOldGlyphId.Count;
            int originalHMetricsCount = originalHmtx.hMetrics.Count;

            // Skapa ny hmtx med rätt storlek
            var newHmtx = originalHmtx.CloneForGlyphCount(finalGlyphCount, context.OriginalFont.MaxpTable.numGlyphs);

            // Fyll i advanceWidth och lsb för varje glyf i subset
            for (int i = 0; i < finalGlyphCount; i++)
            {
                ushort oldGlyphId = context.NewToOldGlyphId[i];

                ushort advanceWidth = originalHmtx.GetAdvanceWidth(oldGlyphId);
                short lsb = originalHmtx.GetLeftSideBearing(oldGlyphId);

                newHmtx.hMetrics[i].advanceWidth = advanceWidth;

                if (i < originalHMetricsCount)
                {
                    newHmtx.hMetrics[i].lsb = lsb;
                }
                else
                {
                    int leftSideBearingIndex = i - originalHMetricsCount;
                    newHmtx.leftSideBearings[leftSideBearingIndex] = lsb;
                }
            }

            // Lägg till i subset-fonten
            newFont.AddOrReplaceTable(newHmtx);
        }
    }
}

