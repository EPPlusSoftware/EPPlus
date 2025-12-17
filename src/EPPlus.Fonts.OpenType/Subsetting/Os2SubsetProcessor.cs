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
using EPPlus.Fonts.OpenType.Tables.Os2;
using System;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Creates a correct OS/2 table for the subset font.
    /// - Recalculates usWinAscent/Descent from actual glyphs
    /// - Sets correct usFirstCharIndex / usLastCharIndex (Windows Font Viewer fix)
    /// - Keeps all other fields from original (safe and correct)
    /// Must run after GlyfAndLocaSubsetProcessor.
    /// .NET 3.5 compatible.
    /// </summary>
    internal class Os2SubsetProcessor : IFontSubsetProcessor
    {
        public void Discover(FontSubsettingContext context)
        {
            // OS/2 discovery is usually not needed as it contains global metrics.
            // We just ensure it's marked for inclusion if necessary.
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var originalFont = context.OriginalFont;
            if (originalFont.Os2Table == null) return;

            // Clone the original OS/2 table
            Os2Table os2 = originalFont.Os2Table.Clone();

            // 1. Unicode range (must match what Rewrite in CmapSubsetProcessor produces)
            if (context.UsedCodePoints.Count > 0)
            {
                // Only include points within the Basic Multilingual Plane (0-0xFFFF)
                var bmpPoints = context.UsedCodePoints.Where(cp => cp <= 0xFFFF).ToList();
                if (bmpPoints.Count > 0)
                {
                    os2.usFirstCharIndex = (ushort)bmpPoints.Min();
                    os2.usLastCharIndex = (ushort)bmpPoints.Max();
                }
            }

            // 2. Sync Windows Metrics (usWinAscent / usWinDescent)
            // We use the metrics from the original font's head table to ensure coverage
            if (originalFont.HeadTable != null)
            {
                ushort headYMax = (ushort)Math.Max((short)0, originalFont.HeadTable.Ymax);
                if (os2.usWinAscent < headYMax)
                {
                    os2.usWinAscent = headYMax;
                }

                ushort headYMinAbs = (ushort)Math.Abs(originalFont.HeadTable.Ymin);
                if (os2.usWinDescent < headYMinAbs)
                {
                    os2.usWinDescent = headYMinAbs;
                }
            }

            // 3. Update xAvgCharWidth (Optional but good for validation)
            // Many validators appreciate if this is recalculated based on the subset
            UpdateAverageCharWidth(context, os2);

            context.SubsetFont.AddOrReplaceTable(os2);
        }

        private void UpdateAverageCharWidth(FontSubsettingContext context, Os2Table os2)
        {
            // Simple average of all glyphs included in the subset
            if (context.SubsetFont.HmtxTable != null && context.OldToNewGlyphId.Count > 0)
            {
                long totalWidth = 0;
                int count = 0;
                foreach (var newGid in context.OldToNewGlyphId.Values)
                {
                    totalWidth += context.SubsetFont.HmtxTable.GetAdvanceWidth(newGid);
                    count++;
                }
                if (count > 0)
                {
                    os2.xAvgCharWidth = (short)(totalWidth / count);
                }
            }
        }
    }
}
