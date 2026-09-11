/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           GSUB Multiple Substitution (Type 2) support
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    /// <summary>
    /// Subsetting handler for GSUB Lookup Type 2 (Multiple Substitution). Unlike Lookup Type 1
    /// (single glyph output) this substitutes ONE input glyph for a SEQUENCE of output glyphs -
    /// every glyph in that sequence is a new reference beyond what the ordinary cmap lookup for
    /// the source text would find, so it must be discovered here the same way a variation
    /// sequence's variant glyph or a ligature's result glyph is.
    /// </summary>
    internal class MultipleSubstHandler : IGsubLookupHandler
    {
        public ushort LookupType => 2;

        public void Discover(FontSubsettingContext context, LookupTable lookup, GsubSubsetProcessor processor)
        {
            // Loop until no new glyphs are added, in case a substituted glyph is itself the
            // trigger for another lookup's substitution (transitive chains), same as
            // SingleSubstHandler does.
            bool addedAny;
            do
            {
                addedAny = false;
                var currentGlyphs = context.IncludedGlyphs.ToArray();

                foreach (var subtable in lookup.SubTables.OfType<MultipleSubstSubTable>())
                {
                    var coveredGids = subtable.Coverage?.GetCoveredGlyphs();
                    if (coveredGids == null) continue;

                    foreach (ushort gid in coveredGids)
                    {
                        if (!context.IncludedGlyphs.Contains(gid))
                            continue;

                        ushort[] sequence = subtable.GetSubstitution(gid);
                        if (sequence == null) continue;

                        foreach (ushort outGid in sequence)
                        {
                            if (!context.IncludedGlyphs.Contains(outGid))
                            {
                                context.IncludedGlyphs.Add(outGid);
                                addedAny = true;
                            }
                        }
                    }
                }
            } while (addedAny);
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            var newLookup = new LookupTable
            {
                LookupType = 2,
                LookupFlag = oldLookup.LookupFlag,
                SubTables = new List<FontTableElement>()
            };

            foreach (var oldSubtable in oldLookup.SubTables.OfType<MultipleSubstSubTable>())
            {
                var coveredGids = oldSubtable.Coverage?.GetCoveredGlyphs();
                if (coveredGids == null) continue;

                var newInputs = new List<ushort>();
                var newSequences = new List<ushort[]>();

                foreach (ushort oldInputGid in coveredGids)
                {
                    // 1. Is the triggering glyph itself part of the subset?
                    if (!context.OldToNewGlyphId.TryGetValue(oldInputGid, out ushort newInputGid))
                        continue;

                    ushort[] oldSequence = oldSubtable.GetSubstitution(oldInputGid);
                    if (oldSequence == null) continue;

                    // 2. Every glyph in the output sequence must also be in the subset - if even
                    // one is missing, the whole substitution is invalid and must be dropped,
                    // rather than emitting a sequence with a dangling .notdef in the middle.
                    var newSequence = new ushort[oldSequence.Length];
                    bool allMapped = true;

                    for (int i = 0; i < oldSequence.Length; i++)
                    {
                        if (!context.OldToNewGlyphId.TryGetValue(oldSequence[i], out ushort newOutGid))
                        {
                            allMapped = false;
                            break;
                        }
                        newSequence[i] = newOutGid;
                    }

                    if (!allMapped) continue;

                    newInputs.Add(newInputGid);
                    newSequences.Add(newSequence);
                }

                if (newInputs.Count == 0) continue;

                // Coverage requires strictly sorted glyph IDs - keep Sequences in the same
                // relative order as the (now sorted) inputs.
                var order = Enumerable.Range(0, newInputs.Count).OrderBy(i => newInputs[i]).ToArray();
                var sortedInputs = order.Select(i => newInputs[i]).ToList();
                var sortedSequences = order.Select(i => newSequences[i]).ToList();

                var newSubtable = new MultipleSubstSubTable
                {
                    SubtableFormat = 1,
                    Coverage = CoverageTableFormat2.CreateCoverageFormat2(sortedInputs),
                    Sequences = sortedSequences
                };
                newLookup.SubTables.Add(newSubtable);
            }

            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }
    }
}