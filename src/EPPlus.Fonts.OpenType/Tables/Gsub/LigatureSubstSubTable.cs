using EPPlus.Fonts.OpenType.Tables.Gsub.Serialization;
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

            // Lista som håller de BaseGlyph ID:n (Gamla ID:n) som överlevde filtreringen
            List<ushort> survivingBaseGlyphs = new List<ushort>();

            // 1. Iterera över befintliga LigatureSets och filtrera dem
            foreach (var kvp in this.LigatureSets)
            {
                ushort oldBaseGlyphId = kvp.Key;
                LigatureSetTable oldLigSet = kvp.Value;

                // BaseGlyph ID:t måste finnas kvar i subsetet.
                ushort newBaseGlyphId;
                bool baseGlyphKept = oldToNewGlyphId.TryGetValue(oldBaseGlyphId, out newBaseGlyphId);

                if (baseGlyphKept)
                {
                    // Skapa den filtrerade LigatureSetTablen
                    LigatureSetTable newLigSet = oldLigSet.CreateSubset(oldToNewGlyphId);

                    if (newLigSet.Ligatures.Count > 0)
                    {
                        // Ligatursetet överlevde och innehåller giltiga ligaturer.

                        // Spara det gamla ID:t för att kunna bygga den nya CoverageTablen
                        survivingBaseGlyphs.Add(oldBaseGlyphId);

                        // Lägg till det nya LigatureSetet (mappat till det GAMLA ID:t)
                        newSubTable.LigatureSets.Add(oldBaseGlyphId, newLigSet);
                    }
                }
            }

            // 2. Återskapa CoverageTablen baserat på de överlevande BaseGlyph ID:na
            if (newSubTable.LigatureSets.Count > 0)
            {
                // För subsetting är Format 1 att föredra.
                CoverageTableFormat1 newCoverage = new CoverageTableFormat1
                {
                    GlyphCount = (ushort)survivingBaseGlyphs.Count,
                    // Sortera listan och använd de nya (ommapade) glyf ID:na
                    // Observera: Det är viktigt att de ommapade ID:na används i GlyfArray
                    GlyphArray = survivingBaseGlyphs.Select(oldId => oldToNewGlyphId[oldId]).OrderBy(g => g).ToArray()
                };
                newSubTable.Coverage = newCoverage;
            }

            return newSubTable;
        }
    }
}
