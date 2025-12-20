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
using EPPlus.Fonts.OpenType.Subsetting;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a Ligature Set table, which contains a list of ligatures 
    /// beginning with a specific first glyph (the base glyph).
    /// </summary>
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
                // 1. Try to map the target glyph (e.g., the "fi" ligature glyph)
                if (!oldToNewGlyphId.TryGetValue(oldLigature.LigatureGlyph, out ushort newTargetGid))
                {
                    continue; // Target glyph is not part of our subset
                }

                // 2. Try to map all subsequent components (e.g., the "i" in "f" + "i")
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
                        break; // A required component is missing; the ligature cannot be formed
                    }
                }

                // 3. If all parts exist in the subset, create a NEW Ligature instance
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
            // Write LigatureCount
            writer.WriteUInt16BigEndian((ushort)this.Ligatures.Count);

            // 1. Calculate and Write Offsets
            // Start offset is after LigatureCount (2 bytes) + all offset entries (2 bytes per ligature)
            int currentOffset = (this.Ligatures.Count * sizeof(ushort)) + sizeof(ushort);

            foreach (var ligature in this.Ligatures)
            {
                writer.WriteUInt16BigEndian((ushort)currentOffset);

                // Calculate size of this LigatureTable to find the next offset:
                // 2 bytes (LigatureGlyph) + 2 bytes (ComponentCount) + (Components.Length * 2 bytes)
                currentOffset += (sizeof(ushort) * 2) + (ligature.Components.Length * sizeof(ushort));
            }

            // 2. Write actual LigatureTable data
            foreach (var ligature in this.Ligatures)
            {
                ligature.Serialize(writer);
            }
        }

        /// <summary>
        /// Rewrites the ligature set based on the subsetting context.
        /// </summary>
        internal LigatureSetTable Rewrite(FontSubsettingContext context)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("  LigatureSetTable.Rewrite: Processing {0} ligatures",
                this.Ligatures.Count));

            LigatureSetTable newSet = new LigatureSetTable();

            foreach (LigatureTable oldLig in this.Ligatures)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("    Ligature: output={0}, components={1}",
                    oldLig.LigatureGlyph,
                    oldLig.Components == null ? "NULL" : string.Join(",", Array.ConvertAll(oldLig.Components, x => x.ToString()))));

                if (!context.OldToNewGlyphId.TryGetValue(oldLig.LigatureGlyph, out ushort newTargetGid))
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("      ❌ Target glyph {0} not in mapping", oldLig.LigatureGlyph));
                    continue;
                }

                System.Diagnostics.Debug.WriteLine(string.Format("      ✅ Target glyph {0} → {1}", oldLig.LigatureGlyph, newTargetGid));

                var newComponents = new List<ushort>();
                bool allComponentsMapped = true;

                foreach (var oldCompGid in oldLig.Components)
                {
                    // ✅ FIX: Ignorera ligatur-komponenter (>= 400), bara mappa base characters
                    if (oldCompGid >= 400)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("        Component {0} is ligature, skipping", oldCompGid));
                        continue; // Skippa ligatur-komponenter
                    }

                    if (context.OldToNewGlyphId.TryGetValue(oldCompGid, out ushort newCompGid))
                    {
                        newComponents.Add(newCompGid);
                        System.Diagnostics.Debug.WriteLine(string.Format("        Component {0} → {1}", oldCompGid, newCompGid));
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("        ❌ Component {0} NOT FOUND", oldCompGid));
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
                    System.Diagnostics.Debug.WriteLine("      ✅ ADDED ligature to new set!");
                }
            }

            System.Diagnostics.Debug.WriteLine(string.Format("  LigatureSetTable.Rewrite: Result has {0} ligatures", newSet.Ligatures.Count));
            return newSet.Ligatures.Count > 0 ? newSet : null;
        }
    }
}