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
            // USHORT SubtableFormat (1)
            writer.WriteUInt16BigEndian(this.SubtableFormat);

            // Placeholder for USHORT CoverageOffset (Vi måste beräkna denna offset)
            long coverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Placeholder

            // USHORT LigSetCount
            writer.WriteUInt16BigEndian((ushort)this.LigatureSets.Count);

            // Placeholder for USHORT[] LigatureSetOffsets (Vi måste beräkna dessa offsets)
            List<long> ligSetOffsetPositions = new List<long>();
            for (int i = 0; i < this.LigatureSets.Count; i++)
            {
                ligSetOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0); // Placeholder
            }

            // --- Skriv ut Sub-tabellerna ---

            // 1. Skriv ut CoverageTable
            long currentOffset = writer.BaseStream.Position;
            long coverageTableOffset = currentOffset;

            if (this.Coverage != null)
            {
                // Skriv ut den relativa offseten till CoverageTable
                ushort relativeCoverageOffset = (ushort)(coverageTableOffset - coverageOffsetPos);
                writer.BaseStream.Seek(coverageOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeCoverageOffset);
                writer.BaseStream.Seek(currentOffset, SeekOrigin.Begin);

                // Serialisera CoverageTable
                if (this.Coverage.CoverageFormat == 1)
                {
                    new CoverageTableFormat1Serializer().Serialize((CoverageTableFormat1)this.Coverage, writer);
                }
                // Lägg till CoverageFormat 2 här om du behöver det, men Format 1 används nu.
            }
            else
            {
                // Om Coverage är null, lämna offseten på 0 (redan gjort)
            }

            // 2. Skriv ut LigatureSetTables
            int ligIndex = 0;

            // De sparade LigatureSets är i ordningen av de överlevande BaseGlyph ID:na (samma ordning som Coverage).
            foreach (var ligSetKvp in this.LigatureSets.OrderBy(kvp => kvp.Value.Ligatures.Min(l => l.LigatureGlyph))) // Använd en stabil sortering
            {
                currentOffset = writer.BaseStream.Position;

                // Skriv ut den relativa offseten till LigatureSetTable
                long ligSetOffsetPos = ligSetOffsetPositions[ligIndex];
                ushort relativeLigSetOffset = (ushort)(currentOffset - ligSetOffsetPos);

                writer.BaseStream.Seek(ligSetOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeLigSetOffset);
                writer.BaseStream.Seek(currentOffset, SeekOrigin.Begin);

                // Serialisera LigatureSetTable
                ligSetKvp.Value.Serialize(writer);
                ligIndex++;
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
