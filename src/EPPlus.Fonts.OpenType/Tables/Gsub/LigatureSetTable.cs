using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class LigatureSetTable : FontTableElement
    {
        /// <summary>
        /// Array of Ligature tables, each corresponding to a substitution sequence
        /// starting with the BaseGlyph (defined by the Coverage table).
        /// </summary>
        public List<LigatureTable> Ligatures { get; set; } = new List<LigatureTable>();

        /// <summary>
        /// Filters and remaps all contained LigatureTable entries.
        /// </summary>
        /// <param name="oldToNewGlyphId">The glyph ID mapping.</param>
        /// <returns>A new, filtered LigatureSetTable containing only valid ligatures.</returns>
        internal LigatureSetTable CreateSubset(Dictionary<ushort, ushort> oldToNewGlyphId)
        {
            LigatureSetTable newSet = new LigatureSetTable();

            foreach (var oldLigature in this.Ligatures)
            {
                // Remap the existing ligature table. The method returns true if the 
                // resulting ligature glyph (output) is kept.
                bool ligatureOutputKept = oldLigature.Remap(oldToNewGlyphId);

                // Now we must check if any component was discarded during Remap 
                // (i.e., mapped to new ID 0, which is .notdef, meaning it was not in the subset).
                bool allComponentsKept = oldLigature.Components.All(gid => gid != 0);

                // If the output ligature is kept AND all input components are kept (i.e., they 
                // were successfully remapped to a non-zero new ID), we keep the ligature.
                if (ligatureOutputKept && allComponentsKept)
                {
                    newSet.Ligatures.Add(oldLigature);
                }
            }

            return newSet;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // LigatureSetTable structure:
            // USHORT LigatureCount
            // USHORT[] LigatureOffsets

            writer.WriteUInt16BigEndian((ushort)this.Ligatures.Count);

            // Calculate offsets for all LigatureTable entries
            // This is complex because we need to write the offsets first, then the actual tables.

            // 1. Calculate and Write Offsets
            int currentOffset = this.Ligatures.Count * sizeof(ushort) + sizeof(ushort); // Start after LigatureCount + all offsets

            foreach (var ligature in this.Ligatures)
            {
                writer.WriteUInt16BigEndian((ushort)currentOffset);
                // Size of LigatureTable: 2 bytes (LigatureGlyph) + 2 bytes (ComponentCount) + Components.Length * 2 bytes
                currentOffset += sizeof(ushort) * 2 + (ligature.Components.Length * sizeof(ushort));
            }

            // 2. Write LigatureTables
            foreach (var ligature in this.Ligatures)
            {
                ligature.Serialize(writer);
            }
        }
    }
}
