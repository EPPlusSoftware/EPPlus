using EPPlus.Fonts.OpenType.Subsetting;
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
                // 1. Försök mappa mål-glyfen (t.ex. "fi")
                if (!oldToNewGlyphId.TryGetValue(oldLigature.LigatureGlyph, out ushort newTargetGid))
                {
                    continue; // Mål-glyfen finns inte i vårt subset
                }

                // 2. Försök mappa alla komponenter (t.ex. "i")
                bool allComponentsMapped = true;
                List<ushort> newComponents = new List<ushort>();

                foreach (var oldCompGid in oldLigature.Components)
                {
                    if (oldToNewGlyphId.TryGetValue(oldCompGid, out ushort newCompGid))
                    {
                        newComponents.Add(newCompGid);
                    }
                    else
                    {
                        allComponentsMapped = false;
                        break; // En komponent saknas, ligaturen kan inte skapas
                    }
                }

                // 3. Om allt finns, skapa en helt NY Ligature-instans
                if (allComponentsMapped)
                {
                    newSet.Ligatures.Add(new LigatureTable
                    {
                        LigatureGlyph = newTargetGid,
                        Components = newComponents.ToArray()
                    });
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

        internal LigatureSetTable Rewrite(FontSubsettingContext context)
        {
            LigatureSetTable newSet = new LigatureSetTable();
            foreach (LigatureTable oldLig in this.Ligatures)
            {
                LigatureTable rewritten = oldLig.CloneAndRewrite(context);
                if (rewritten != null)
                {
                    newSet.Ligatures.Add(rewritten);
                }
            }
            return newSet.Ligatures.Count > 0 ? newSet : null;
        }
    }
}
