using EPPlus.Fonts.OpenType.Subsetting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    internal class SingleSubstHandler : IGsubLookupHandler
    {
        public ushort LookupType => 1;

        public void Discover(FontSubsettingContext context, LookupTable lookup)
        {
            // Vi behöver veta om vi la till något för att stödja den iterativa loopen i processorn
            // Men interfacet är void, så vi låter context.IncludedGlyphs hantera tillståndet.
            var currentGlyphs = context.IncludedGlyphs.ToArray();
            foreach (var subtable in lookup.SubTables.OfType<SingleSubstSubTable>())
            {
                foreach (ushort gid in currentGlyphs)
                {
                    ushort substitute = subtable.GetSubstitution(gid);
                    if (substitute != 0 && !context.IncludedGlyphs.Contains(substitute))
                    {
                        context.IncludedGlyphs.Add(substitute);
                        // Eftersom GsubSubsetProcessor kör en do-while loop 
                        // kommer den upptäcka att listan har vuxit.
                    }
                }
            }
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            var newLookup = new LookupTable { LookupType = 1, LookupFlag = oldLookup.LookupFlag, SubTables = new List<FontTableElement>() };
            foreach (var subtable in oldLookup.SubTables.OfType<SingleSubstSubTable>())
            {
                var rewritten = subtable.Rewrite(context);
                if (rewritten != null) newLookup.SubTables.Add(rewritten);
            }
            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }
    }
}
