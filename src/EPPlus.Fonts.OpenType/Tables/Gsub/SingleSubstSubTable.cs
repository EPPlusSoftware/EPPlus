using EPPlus.Fonts.OpenType.Subsetting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public abstract class SingleSubstSubTable : FontTableElement
    {
        public ushort SubtableFormat { get; set; }
        public CoverageTable Coverage { get; set; }

        // Gemensam metod för att få ut substitutionen för en Base Glyph ID.
        // Måste implementeras i varje format.
        public abstract ushort GetSubstitution(ushort baseGlyphId);

        public SingleSubstSubTable Rewrite(FontSubsettingContext context)
        {
            // En lista för att temporärt hålla våra par av (Nytt Input GID, Nytt Output GID)
            List<GsubRewriteEntry> validMappings = new List<GsubRewriteEntry>();

            // Hämta alla Glyph IDs som denna tabell hanterar
            ushort[] oldInputGlyphs = this.Coverage.GetCoveredGlyphs();

            foreach (ushort oldInputGid in oldInputGlyphs)
            {
                ushort newInputGid;
                // 1. Ska tecknet som triggar substitutionen vara med?
                if (context.GlyphIdMap.TryGetValue(oldInputGid, out newInputGid))
                {
                    ushort oldOutputGid = GetSubstitution(oldInputGid);
                    ushort newOutputGid;

                    // 2. Ska tecknet som man byter TILL också vara med?
                    if (context.GlyphIdMap.TryGetValue(oldOutputGid, out newOutputGid))
                    {
                        GsubRewriteEntry entry;
                        entry.NewInput = newInputGid;
                        entry.NewOutput = newOutputGid;
                        validMappings.Add(entry);
                    }
                }
            }

            if (validMappings.Count == 0) return null;

            // Sortera efter NewInput - ett krav för CoverageTable
            validMappings.Sort(delegate (GsubRewriteEntry a, GsubRewriteEntry b) {
                return a.NewInput.CompareTo(b.NewInput);
            });

            // Skapa den nya tabellen
            SingleSubstSubTableFormat2 newTable = new SingleSubstSubTableFormat2();
            List<ushort> newInputs = new List<ushort>();
            newTable.SubstituteGlyphIDs = new ushort[validMappings.Count];

            for (int i = 0; i < validMappings.Count; i++)
            {
                newInputs.Add(validMappings[i].NewInput);
                newTable.SubstituteGlyphIDs[i] = validMappings[i].NewOutput;
            }

            newTable.Coverage = CoverageTableFormat2.CreateCoverageFormat2(newInputs);
            newTable.GlyphCount = (ushort)validMappings.Count;

            return newTable;
        }

        private struct GsubRewriteEntry
        {
            public ushort NewInput;
            public ushort NewOutput;
        }
    }
}
