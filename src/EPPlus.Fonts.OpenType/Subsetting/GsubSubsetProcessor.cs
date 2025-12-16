using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class GsubSubsetProcessor : IFontSubsetProcessor
    {
        private Dictionary<int, int> _oldToNewLookupIndex = new Dictionary<int, int>();
        private Dictionary<int, int> _oldToNewFeatureIndex = new Dictionary<int, int>();

        public void Process(FontSubsettingContext context)
        {
            var gsub = context.OriginalFont.GsubTable;
            if (gsub == null || gsub.LookupList == null) return;

            // Phase 1: Discovery
            // Vi måste hitta alla glyfer som GSUB kan tänkas skapa (t.ex. ligaturer)
            // och lägga till dem i context.IncludedGlyphs så att de får ett nytt ID.
            bool glyphsAdded;
            do
            {
                glyphsAdded = false;
                ushort[] currentGlyphs = context.IncludedGlyphs.ToArray();

                foreach (var lookup in gsub.LookupList.Lookups)
                {
                    foreach (var subtable in lookup.SubTables)
                    {
                        if (subtable is SingleSubstSubTable single)
                        {
                            foreach (ushort gid in currentGlyphs)
                            {
                                ushort substitute = single.GetSubstitution(gid);
                                if (substitute != 0 && !context.IncludedGlyphs.Contains(substitute))
                                {
                                    context.IncludedGlyphs.Add(substitute);
                                    glyphsAdded = true;
                                }
                            }
                        }
                        else if (subtable is LigatureSubstSubTable lig)
                        {
                            if (DiscoverLigatures(context, lig))
                            {
                                glyphsAdded = true;
                            }
                        }
                    }
                }
            } while (glyphsAdded);
        }

        private bool DiscoverLigatures(FontSubsettingContext context, LigatureSubstSubTable subtable)
        {
            bool addedAny = false;
            if (subtable.LigatureSets == null) return false;

            foreach (var kvp in subtable.LigatureSets)
            {
                ushort firstGlyph = kvp.Key;
                if (!context.IncludedGlyphs.Contains(firstGlyph)) continue;

                foreach (var lig in kvp.Value.Ligatures)
                {
                    bool allComponentsPresent = true;
                    if (lig.Components != null)
                    {
                        foreach (ushort compGid in lig.Components)
                        {
                            if (!context.IncludedGlyphs.Contains(compGid))
                            {
                                allComponentsPresent = false;
                                break;
                            }
                        }
                    }

                    if (allComponentsPresent && !context.IncludedGlyphs.Contains(lig.LigatureGlyph))
                    {
                        context.IncludedGlyphs.Add(lig.LigatureGlyph);
                        addedAny = true;
                    }
                }
            }
            return addedAny;
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var oldGsub = context.OriginalFont.GsubTable;
            if (oldGsub == null) return;

            var newGsub = new GsubTable();
            _oldToNewLookupIndex.Clear();
            _oldToNewFeatureIndex.Clear();

            newGsub.LookupList = RemapLookupList(context, oldGsub.LookupList);
            newGsub.FeatureList = RemapFeatureList(context, oldGsub.FeatureList);
            newGsub.ScriptList = RemapScriptList(context, oldGsub.ScriptList);

            context.SubsetFont.AddOrReplaceTable(newGsub);
        }

        private LookupListTable RemapLookupList(FontSubsettingContext context, LookupListTable oldLookupList)
        {
            var newList = new LookupListTable();
            if (oldLookupList == null) return newList;

            for (int i = 0; i < oldLookupList.Lookups.Count; i++)
            {
                var oldLookup = oldLookupList.Lookups[i];
                var newLookup = new LookupTable
                {
                    LookupType = oldLookup.LookupType,
                    LookupFlag = oldLookup.LookupFlag
                };

                foreach (var oldSub in oldLookup.SubTables)
                {
                    var remappedSub = RemapSubtable(context, oldSub);
                    if (remappedSub != null)
                        newLookup.SubTables.Add(remappedSub);
                }

                if (newLookup.SubTables.Count > 0)
                {
                    _oldToNewLookupIndex[i] = newList.Lookups.Count;
                    newList.Lookups.Add(newLookup);
                }
            }
            return newList;
        }

        private FontTableElement RemapSubtable(FontSubsettingContext context, FontTableElement oldSub)
        {
            if (oldSub is SingleSubstSubTableFormat1 f1) return RemapSingleSubstFormat1(context, f1);
            if (oldSub is SingleSubstSubTableFormat2 f2) return RemapSingleSubstFormat2(context, f2);
            if (oldSub is LigatureSubstSubTable lig) return RemapLigatureSubst(context, lig);
            return null;
        }

        private FontTableElement RemapSingleSubstFormat1(FontSubsettingContext context, SingleSubstSubTableFormat1 oldSub)
        {
            // Konvertera Format 1 -> Format 2 för säkerhets skull vid subsetting
            var substitutes = new List<ushort>();
            var activeOldGids = new List<ushort>();
            ushort[] covered = oldSub.Coverage.GetCoveredGlyphs();

            foreach (var oldGid in covered)
            {
                if (context.IncludedGlyphs.Contains(oldGid))
                {
                    ushort oldSubstitute = (ushort)(oldGid + oldSub.DeltaGlyphID);
                    if (context.IncludedGlyphs.Contains(oldSubstitute))
                    {
                        substitutes.Add(GetNewId(context, oldSubstitute));
                        activeOldGids.Add(oldGid);
                    }
                }
            }

            if (activeOldGids.Count == 0) return null;

            return new SingleSubstSubTableFormat2
            {
                SubtableFormat = 2,
                SubstituteGlyphIDs = substitutes.ToArray(),
                GlyphCount = (ushort)substitutes.Count,
                Coverage = CreateNewCoverage(context, activeOldGids)
            };
        }

        private FontTableElement RemapSingleSubstFormat2(FontSubsettingContext context, SingleSubstSubTableFormat2 oldSub)
        {
            var substitutes = new List<ushort>();
            var activeOldGids = new List<ushort>();
            ushort[] covered = oldSub.Coverage.GetCoveredGlyphs();

            for (int i = 0; i < covered.Length; i++)
            {
                ushort oldGid = covered[i];
                if (context.IncludedGlyphs.Contains(oldGid))
                {
                    ushort oldSubst = oldSub.SubstituteGlyphIDs[i];
                    if (context.IncludedGlyphs.Contains(oldSubst))
                    {
                        substitutes.Add(GetNewId(context, oldSubst));
                        activeOldGids.Add(oldGid);
                    }
                }
            }

            if (activeOldGids.Count == 0) return null;

            return new SingleSubstSubTableFormat2
            {
                SubtableFormat = 2,
                SubstituteGlyphIDs = substitutes.ToArray(),
                GlyphCount = (ushort)substitutes.Count,
                Coverage = CreateNewCoverage(context, activeOldGids)
            };
        }

        private FontTableElement RemapLigatureSubst(FontSubsettingContext context, LigatureSubstSubTable oldSub)
        {
            var newLigSets = new Dictionary<ushort, LigatureSetTable>();
            var activeOldFirstGlyphs = new List<ushort>();

            foreach (var kvp in oldSub.LigatureSets)
            {
                ushort oldFirst = kvp.Key;
                if (!context.IncludedGlyphs.Contains(oldFirst)) continue;

                var newSet = new LigatureSetTable();
                foreach (var oldLig in kvp.Value.Ligatures)
                {
                    bool componentsValid = oldLig.Components == null || oldLig.Components.All(c => context.IncludedGlyphs.Contains(c));
                    if (componentsValid && context.IncludedGlyphs.Contains(oldLig.LigatureGlyph))
                    {
                        newSet.Ligatures.Add(new LigatureTable
                        {
                            LigatureGlyph = GetNewId(context, oldLig.LigatureGlyph),
                            Components = oldLig.Components?.Select(c => GetNewId(context, c)).ToArray()
                        });
                    }
                }

                if (newSet.Ligatures.Count > 0)
                {
                    ushort newFirst = GetNewId(context, oldFirst);
                    newLigSets[newFirst] = newSet;
                    activeOldFirstGlyphs.Add(oldFirst);
                }
            }

            if (newLigSets.Count == 0) return null;

            return new LigatureSubstSubTable
            {
                SubtableFormat = 1,
                LigatureSets = newLigSets,
                Coverage = CreateNewCoverage(context, activeOldFirstGlyphs)
            };
        }

        private CoverageTable CreateNewCoverage(FontSubsettingContext context, List<ushort> oldGids)
        {
            var newIds = oldGids.Select(id => GetNewId(context, id)).Distinct().ToList();
            newIds.Sort();
            return new CoverageTableFormat1
            {
                CoverageFormat = 1,
                GlyphArray = newIds.ToArray(),
                GlyphCount = (ushort)newIds.Count
            };
        }

        private ushort GetNewId(FontSubsettingContext context, ushort oldId)
        {
            return context.OldToNewGlyphId.TryGetValue(oldId, out var newId) ? newId : (ushort)0;
        }

        // ... RemapFeatureList, RemapScriptList och RemapLangSys behålls som de var tidigare ...
        private FeatureListTable RemapFeatureList(FontSubsettingContext context, FeatureListTable oldList)
        {
            var newList = new FeatureListTable();
            if (oldList == null) return newList;
            var tempEntries = new List<FeatureMappingEntry>();

            for (int i = 0; i < oldList.FeatureRecords.Count; i++)
            {
                var oldRecord = oldList.FeatureRecords[i];
                var newIndices = oldRecord.FeatureTable.LookupListIndices
                    .Where(idx => _oldToNewLookupIndex.ContainsKey(idx))
                    .Select(idx => (ushort)_oldToNewLookupIndex[idx])
                    .ToArray();

                if (newIndices.Length > 0)
                {
                    tempEntries.Add(new FeatureMappingEntry
                    {
                        OldIndex = i,
                        Record = new FeatureRecord
                        {
                            FeatureTag = oldRecord.FeatureTag,
                            FeatureTable = new FeatureTable { LookupListIndices = newIndices, FeatureParams = oldRecord.FeatureTable.FeatureParams }
                        }
                    });
                }
            }

            tempEntries.Sort((a, b) => string.CompareOrdinal(a.Record.FeatureTag?.Value, b.Record.FeatureTag?.Value));
            for (int i = 0; i < tempEntries.Count; i++)
            {
                _oldToNewFeatureIndex[tempEntries[i].OldIndex] = i;
                newList.FeatureRecords.Add(tempEntries[i].Record);
            }
            return newList;
        }

        private ScriptListTable RemapScriptList(FontSubsettingContext context, ScriptListTable oldList)
        {
            var newList = new ScriptListTable();
            if (oldList == null) return newList;
            foreach (var oldRec in oldList.ScriptRecords)
            {
                var newScript = new ScriptTable();
                if (oldRec.ScriptTable.DefaultLangSys != null)
                    newScript.DefaultLangSys = RemapLangSys(oldRec.ScriptTable.DefaultLangSys);

                foreach (var lsr in oldRec.ScriptTable.LangSysRecords)
                    newScript.LangSysRecords.Add(new LangSysRecord { LangSysTag = lsr.LangSysTag, LangSysTable = RemapLangSys(lsr.LangSysTable) });

                newList.ScriptRecords.Add(new ScriptRecord { ScriptTag = oldRec.ScriptTag, ScriptTable = newScript });
            }
            newList.ScriptRecords.Sort((a, b) => string.CompareOrdinal(a.ScriptTag?.Value, b.ScriptTag?.Value));
            return newList;
        }

        private LangSysTable RemapLangSys(LangSysTable oldLang)
        {
            var newLang = new LangSysTable { LookupOrder = oldLang.LookupOrder, RequiredFeatureIndex = 0xFFFF };
            if (oldLang.RequiredFeatureIndex != 0xFFFF && _oldToNewFeatureIndex.TryGetValue(oldLang.RequiredFeatureIndex, out var newReq))
                newLang.RequiredFeatureIndex = (ushort)newReq;

            newLang.FeatureIndices = oldLang.FeatureIndices
                .Where(idx => _oldToNewFeatureIndex.ContainsKey(idx))
                .Select(idx => (ushort)_oldToNewFeatureIndex[idx])
                .ToArray();
            newLang.FeatureIndexCount = (ushort)newLang.FeatureIndices.Length;
            return newLang;
        }

        private class FeatureMappingEntry { public int OldIndex; public FeatureRecord Record; }
    }
}