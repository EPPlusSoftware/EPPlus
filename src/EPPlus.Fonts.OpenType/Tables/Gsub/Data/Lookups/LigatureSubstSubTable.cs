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
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a Ligature Substitution Subtable (Lookup Type 4, Format 1).
    /// This table maps a Base Glyph (via the Coverage table) to a LigatureSetTable.
    /// </summary>
    public class LigatureSubstSubTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the format identifier for the subtable (should be 1).
        /// </summary>
        public ushort SubtableFormat { get; set; }

        /// <summary>
        /// Gets or sets the Coverage table which defines the set of initial glyphs (Base Glyphs) 
        /// that start the ligature sequence.
        /// </summary>
        public CoverageTable Coverage { get; set; }

        /// <summary>
        /// Gets or sets a dictionary mapping the Base Glyph ID (from the Coverage table) 
        /// to the corresponding Ligature Set.
        /// </summary>
        public Dictionary<ushort, LigatureSetTable> LigatureSets { get; set; } = new Dictionary<ushort, LigatureSetTable>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. Store start position for relative offset calculations
            long subTableStart = writer.BaseStream.Position;

            // 2. Write Header
            writer.WriteUInt16BigEndian(this.SubtableFormat); // Usually 1

            // 3. Placeholder for CoverageOffset
            long covOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 4. Write number of LigatureSets
            writer.WriteUInt16BigEndian((ushort)this.LigatureSets.Count);

            // 5. Write placeholders for each LigatureSetOffset
            // These must be written in the exact same order as the glyphs in the CoverageTable
            List<long> ligSetOffsetPositions = new List<long>();
            for (int i = 0; i < this.LigatureSets.Count; i++)
            {
                ligSetOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // --- DATA SERIALIZATION ---

            // 6. Serialize CoverageTable and backfill its offset
            if (this.Coverage != null)
            {
                this.WriteRelativeOffset(writer, subTableStart, covOffsetPos);
                this.Coverage.Serialize(writer);
            }

            // 7. Serialize LigatureSets (Order must match CoverageTable indices)
            ushort[] coveredGlyphs = this.Coverage.GetCoveredGlyphs();
            for (int i = 0; i < coveredGlyphs.Length; i++)
            {
                ushort baseGlyphId = coveredGlyphs[i];

                if (this.LigatureSets.TryGetValue(baseGlyphId, out var ligSet))
                {
                    // Update the offset for this specific set in the array from step 5
                    this.WriteRelativeOffset(writer, subTableStart, ligSetOffsetPositions[i]);

                    // Serialize the LigatureSetTable (which handles its own internal offsets)
                    ligSet.Serialize(writer);
                }
            }
        }

        /// <summary>
        /// Filters the contained LigatureSets based on the subset mapping, 
        /// removes obsolete ligatures, and reconstructs the CoverageTable.
        /// </summary>
        internal LigatureSubstSubTable CreateSubset(Dictionary<ushort, ushort> oldToNewGlyphId)
        {
            LigatureSubstSubTable newSubTable = new LigatureSubstSubTable { SubtableFormat = this.SubtableFormat };
            List<ushort> newBaseGlyphs = new List<ushort>();

            // 1. Iterate over existing LigatureSets
            foreach (var kvp in this.LigatureSets)
            {
                ushort oldBaseGlyphId = kvp.Key;
                LigatureSetTable oldLigSet = kvp.Value;

                // Check if the starting glyph (e.g., 'f') exists in the subset
                if (oldToNewGlyphId.TryGetValue(oldBaseGlyphId, out ushort newBaseGlyphId))
                {
                    // Create filtered LigatureSet (remapping all internal component GIDs)
                    LigatureSetTable newLigSet = oldLigSet.CreateSubset(oldToNewGlyphId);

                    if (newLigSet != null && newLigSet.Ligatures.Count > 0)
                    {
                        // Store using the NEW Glyph ID as the key
                        newSubTable.LigatureSets[newBaseGlyphId] = newLigSet;
                        newBaseGlyphs.Add(newBaseGlyphId);
                    }
                }
            }

            // 2. Reconstruct the Coverage Table
            if (newSubTable.LigatureSets.Count > 0)
            {
                // OpenType requires the Coverage table glyphs to be sorted numerically
                newBaseGlyphs.Sort();

                newSubTable.Coverage = new CoverageTableFormat1
                {
                    GlyphCount = (ushort)newBaseGlyphs.Count,
                    GlyphArray = newBaseGlyphs.ToArray()
                };
            }

            return newSubTable.LigatureSets.Count > 0 ? newSubTable : null;
        }

        /// <summary>
        /// Rewrites the subtable using the provided subsetting context.
        /// </summary>
        public LigatureSubstSubTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            var newSubTable = new LigatureSubstSubTable();
            newSubTable.SubtableFormat = this.SubtableFormat;
            newSubTable.LigatureSets = new Dictionary<ushort, LigatureSetTable>();

            foreach (var oldSet in this.LigatureSets)
            {
                // 1. Map the start glyph (e.g., 'f')
                if (!context.OldToNewGlyphId.TryGetValue(oldSet.Key, out ushort newFirstGid))
                    continue;

                var newSet = new LigatureSetTable();
                newSet.Ligatures = new List<LigatureTable>();

                foreach (var oldLig in oldSet.Value.Ligatures)
                {
                    // 2. Map the target ligature glyph (e.g., 'fi')
                    if (!context.OldToNewGlyphId.TryGetValue(oldLig.LigatureGlyph, out ushort newTargetGid))
                        continue;

                    // 3. Map all components (e.g., 'i' in "fi")
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

            if (newSubTable.LigatureSets.Count > 0)
            {
                var newBaseGlyphs = new List<ushort>(newSubTable.LigatureSets.Keys);
                newBaseGlyphs.Sort(); // OpenType kräver sorterad Coverage

                newSubTable.Coverage = new CoverageTableFormat1
                {
                    GlyphCount = (ushort)newBaseGlyphs.Count,
                    GlyphArray = newBaseGlyphs.ToArray()
                };
            }

            return newSubTable.LigatureSets.Count > 0 ? newSubTable : null;
        }
    }
}