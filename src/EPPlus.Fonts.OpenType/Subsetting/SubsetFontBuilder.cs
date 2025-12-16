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
using EPPlus.Fonts.OpenType.Subsetting.Processors;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class SubsetFontBuilder
    {

        private static IEnumerable<IFontSubsetProcessor> Processors => new List<IFontSubsetProcessor>
        {
            // 1. DISCOVERY - Dessa måste köra först för att fylla IncludedGlyphs
            new CmapSubsetProcessor(),  // Hittar GIDs för dina tecken
            new GsubSubsetProcessor(),  // Hittar GIDs för substitutioner/ligaturer
    
            // 2. DATA EXTRACTION - Nu när vi vet ALLA GIDs som behövs, hämta datan
            new GlyfAndLocaSubsetProcessor(), 
    
            // 3. METADATA & ÖVRIGT
            new MaxpSubsetProcessor(),
            new HeadSubsetProcessor(),
            new NameSubsetProcessor(),
            new HheaSubsetProcessor(),
            new HmtxSubsetProcessor(),
            new Os2SubsetProcessor(),
            new PostSubsetProcessor(),
            new KernSubsetProcessor()
        };

        public OpenTypeFont CreateSubset(OpenTypeFont originalFont, IEnumerable<int> unicodeChars)
        {
            var context = new FontSubsettingContext(originalFont, unicodeChars);
            var processors = Processors; // Hämta listan en gång

            // Steg 1: Discovery Phase
            // Alla processorer (inkl. GSUB) hittar vilka glyfer som behövs.
            foreach (var processor in processors)
            {
                processor.Process(context);
            }

            // Steg 2: Skapa Glyph ID-mappningen (Viktigt!)
            // Här går vi från IncludedGlyphs (HashSet) till OldToNewGlyphId (Dictionary)
            BuildGlyphMapping(context);

            // Steg 3: Rewrite Phase
            // Nu när context.OldToNewGlyphId är populerad kan GSUB och andra tabeller skrivas om.
            foreach (var processor in processors)
            {
                // Vi kan lägga till en check här, eller låta GsubSubsetProcessor 
                // internt anropa Rewrite från sin Process-metod (se nästa steg).
                if (processor is GsubSubsetProcessor gsubProcessor)
                {
                    gsubProcessor.Rewrite(context);
                }
                else if(processor is Os2SubsetProcessor os2Processor)
                {
                    os2Processor.Rewrite(context);
                }
                else if(processor is CmapSubsetProcessor cmapProcessor)
                {
                    cmapProcessor.Rewrite(context);
                }
                // hmtx, post, cmap etc. kan också behöva anropas här om de inte 
                // redan sköts inuti sin Process.
            }

            // 9. Debug-info
            context.SubsetFont.UsedCodePointsForSubset = new List<uint>(context.UsedCodePoints);

            return context.SubsetFont;
        }

        private void BuildGlyphMapping(FontSubsettingContext context)
        {
            // Sortera för att få deterministiska Glyph IDs i den nya fonten
            List<ushort> sortedGlyphs = new List<ushort>(context.IncludedGlyphs);
            sortedGlyphs.Sort();

            for (ushort newId = 0; newId < sortedGlyphs.Count; newId++)
            {
                ushort oldId = sortedGlyphs[newId];
                context.OldToNewGlyphId[oldId] = newId;
                context.NewToOldGlyphId.Add(oldId);
            }
        }
    }
}