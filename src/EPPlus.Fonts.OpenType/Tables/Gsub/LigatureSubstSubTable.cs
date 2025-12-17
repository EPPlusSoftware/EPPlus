using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    /// <summary>
    /// Represents a Ligature Substitution Subtable (Type 4, Format 1).
    /// This table maps a Base Glyph (via the Coverage table) to a LigatureSetTable.
    /// </summary>
    public class LigatureSubstSubTable : FontTableElement
    {
        /// <summary>
        /// Format identifier for the subtable (should be 1).
        /// </summary>
        public ushort SubtableFormat { get; set; }

        /// <summary>
        /// The Coverage table which defines the set of initial glyphs (Base Glyphs) 
        /// that start the ligature sequence.
        /// </summary>
        public CoverageTable Coverage { get; set; }

        /// <summary>
        /// A dictionary mapping the Base Glyph ID (from Coverage table) 
        /// to the corresponding Ligature Set.
        /// </summary>
        public Dictionary<ushort, LigatureSetTable> LigatureSets { get; set; } = new Dictionary<ushort, LigatureSetTable>();

        // We will implement Serialize later
        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. Spara startposition för hela subtabellen
            long subTableStart = writer.BaseStream.Position;

            // 2. Skriv Header
            writer.WriteUInt16BigEndian(this.SubtableFormat); // Borde vara 1

            // 3. Reservera plats för CoverageOffset
            long covOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 4. Skriv antal LigatureSets
            writer.WriteUInt16BigEndian((ushort)this.LigatureSets.Count);

            // 5. Skriv placeholders för varje LigatureSetOffset
            // Dessa måste skrivas i samma ordning som glyferna i CoverageTable
            List<long> ligSetOffsetPositions = new List<long>();
            for (int i = 0; i < this.LigatureSets.Count; i++)
            {
                ligSetOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // --- SKRIV DATA-TABELLER ---

            // 6. Serialisera CoverageTable och uppdatera dess offset
            if (this.Coverage != null)
            {
                this.WriteRelativeOffset(writer, subTableStart, covOffsetPos);
                this.Coverage.Serialize(writer);
            }

            // 7. Serialisera LigatureSets (Viktigt: Måste följa Coverage-ordningen!)
            ushort[] coveredGlyphs = this.Coverage.GetCoveredGlyphs();
            for (int i = 0; i < coveredGlyphs.Length; i++)
            {
                ushort baseGlyphId = coveredGlyphs[i];

                if (this.LigatureSets.TryGetValue(baseGlyphId, out var ligSet))
                {
                    // Uppdatera offseten för detta specifika set i arrayen vi skrev i steg 5
                    this.WriteRelativeOffset(writer, subTableStart, ligSetOffsetPositions[i]);

                    // Låt LigatureSetTable skriva sig själv (inklusive sina egna LigatureOffsets)
                    ligSet.Serialize(writer);
                }
            }
        }

        /// <summary>
        /// Filters the contained LigatureSets based on the subset mapping, 
        /// removes obsolete ligatures, and reconstructs the CoverageTable.
        /// </summary>
        /// <param name="oldToNewGlyphId">The glyph ID mapping.</param>
        /// <returns>A new LigatureSubstSubTable object ready for serialization.</returns>
        internal LigatureSubstSubTable CreateSubset(Dictionary<ushort, ushort> oldToNewGlyphId)
        {
            LigatureSubstSubTable newSubTable = new LigatureSubstSubTable { SubtableFormat = this.SubtableFormat };

            // Vi använder en temporär lista för att hålla reda på vilka nya ID:n vi skapar ligaturer för
            // för att sedan kunna bygga en korrekt CoverageTable.
            List<ushort> newBaseGlyphs = new List<ushort>();

            // 1. Iterera över befintliga LigatureSets
            foreach (var kvp in this.LigatureSets)
            {
                ushort oldBaseGlyphId = kvp.Key;
                LigatureSetTable oldLigSet = kvp.Value;

                // Kontrollera om start-glyfen (t.ex. 'f') ska vara med i subsetet
                if (oldToNewGlyphId.TryGetValue(oldBaseGlyphId, out ushort newBaseGlyphId))
                {
                    // Skapa det filtrerade LigatureSetet (här inne måste också alla GIDs mappas om!)
                    LigatureSetTable newLigSet = oldLigSet.CreateSubset(oldToNewGlyphId);

                    if (newLigSet != null && newLigSet.Ligatures.Count > 0)
                    {
                        // VIKTIGT: Vi sparar nu med det NYA ID:t som nyckel!
                        // Detta gör att Serialize-metodens TryGetValue kommer fungera.
                        newSubTable.LigatureSets[newBaseGlyphId] = newLigSet;
                        newBaseGlyphs.Add(newBaseGlyphId);
                    }
                }
            }

            // 2. Återskapa CoverageTablen
            if (newSubTable.LigatureSets.Count > 0)
            {
                // Sortera de nya ID-värdena. OpenType kräver att Coverage-tabellen är sorterad.
                newBaseGlyphs.Sort();

                newSubTable.Coverage = new CoverageTableFormat1
                {
                    GlyphCount = (ushort)newBaseGlyphs.Count,
                    GlyphArray = newBaseGlyphs.ToArray()
                };
            }

            return newSubTable.LigatureSets.Count > 0 ? newSubTable : null;
        }

        public LigatureSubstSubTable Rewrite(FontSubsettingContext context)
        {
            var newSubTable = new LigatureSubstSubTable();
            newSubTable.LigatureSets = new Dictionary<ushort, LigatureSetTable>();

            foreach (var oldSet in this.LigatureSets)
            {
                // 1. Mappa om Start-glyf (t.ex. 'f')
                // Om 'f' inte finns i vårt subset, hoppa över hela setet
                if (!context.OldToNewGlyphId.TryGetValue(oldSet.Key, out ushort newFirstGid))
                    continue;

                var newSet = new LigatureSetTable();
                newSet.Ligatures = new List<LigatureTable>();

                foreach (var oldLig in oldSet.Value.Ligatures)
                {
                    // 2. Mappa om Mål-glyf (t.ex. 'fi')
                    if (!context.OldToNewGlyphId.TryGetValue(oldLig.LigatureGlyph, out ushort newTargetGid))
                        continue;

                    // 3. Mappa om alla komponenter (t.ex. 'i' i "fi", eller 'f','l' i "ffl")
                    var newComponents = new List<ushort>();
                    bool allComponentsMapped = true;

                    foreach (var oldCompGid in oldLig.Components)
                    {
                        if (context.OldToNewGlyphId.TryGetValue(oldCompGid, out ushort newCompGid))
                        {
                            newComponents.Add(newCompGid);
                        }
                        else
                        {
                            allComponentsMapped = false;
                            break;
                        }
                    }

                    if (allComponentsMapped)
                    {
                        newSet.Ligatures.Add(new LigatureTable
                        {
                            LigatureGlyph = newTargetGid,
                            Components = newComponents.ToArray()
                        });
                    }
                }

                if (newSet.Ligatures.Count > 0)
                {
                    newSubTable.LigatureSets[newFirstGid] = newSet;
                }
            }

            return newSubTable.LigatureSets.Count > 0 ? newSubTable : null;
        }
    }
}
