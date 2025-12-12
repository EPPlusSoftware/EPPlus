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
        public void Process(FontSubsettingContext context)
        {
            var originalFont = context.OriginalFont;
            var subsetFont = context.SubsetFont;

            if (originalFont.Os2Table == null)
                return; // inget OS/2 i originalet

            Os2Table os2 = originalFont.Os2Table.Clone();

            // --------------------------------------------------------------------
            // 1. Set correct usFirstCharIndex / usLastCharIndex
            // → Ensures Windows Font Viewer shows the preview text instead of a blank window
            // --------------------------------------------------------------------
            if (context.UsedCodePoints.Count > 0)
            {
                uint first = context.UsedCodePoints.Min();
                uint last = context.UsedCodePoints.Max();

                os2.usFirstCharIndex = (ushort)Math.Max(32, Math.Min(first, 0xFFFF));
                os2.usLastCharIndex = (ushort)Math.Min(last, 0xFFFF);
            }
            else
            {
                os2.usFirstCharIndex = 32;
                os2.usLastCharIndex = 32;
            }

            // --------------------------------------------------------------------
            // 2. Recalculate usWinAscent / usWinDescent based on actual glyphs in the subset
            // → Prevents clipped accents in Windows applications (especially important for åäö, č, đ, etc.)
            // --------------------------------------------------------------------
            short maxAscent = short.MinValue;
            short minDescent = short.MaxValue;

            foreach (var glyph in subsetFont.GlyfTable.Glyphs)
            {
                if (glyph?.Header != null)
                {
                    if (glyph.Header.yMax > maxAscent) maxAscent = glyph.Header.yMax;
                    if (glyph.Header.yMin < minDescent) minDescent = glyph.Header.yMin;
                }
            }

            // Uppdatera bara om subset har högre/lägre värden än original
            if (maxAscent > 0 && maxAscent > os2.usWinAscent)
                os2.usWinAscent = (ushort)maxAscent;

            if (minDescent < 0 && (ushort)(-minDescent) > os2.usWinDescent)
                os2.usWinDescent = (ushort)(-minDescent);

            // --------------------------------------------------------------------
            // 3. Synchronize hhea table with OS/2 typo metrics (best practice)
            // --------------------------------------------------------------------
            if (subsetFont.HheaTable != null)
            {
                subsetFont.HheaTable.ascender = os2.sTypoAscender;
                subsetFont.HheaTable.descender = os2.sTypoDescender;
                subsetFont.HheaTable.lineGap = os2.sTypoLineGap;
            }

            // --------------------------------------------------------------------
            // 4. Store the updated OS/2 table in the subset font
            // --------------------------------------------------------------------
            subsetFont.AddOrReplaceTable(os2);
        }
    }
}
