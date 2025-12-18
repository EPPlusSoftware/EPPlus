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
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data;
using EPPlus.Fonts.OpenType.Tables.Gsub.Handlers;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class GsubSubsetProcessor : IFontSubsetProcessor
    {
        private Dictionary<int, int> _oldToNewLookupIndex = new Dictionary<int, int>();
        private Dictionary<int, int> _oldToNewFeatureIndex = new Dictionary<int, int>();
        private readonly Dictionary<ushort, IGsubLookupHandler> _handlers;

        public GsubSubsetProcessor()
        {
            var handlers = new IGsubLookupHandler[]
            {
                new SingleSubstHandler(),
                new LigatureSubstHandler(),
            };
            _handlers = handlers.ToDictionary(h => h.LookupType);
        }

        public void Discover(FontSubsettingContext context)
        {
            var gsub = context.OriginalFont.GsubTable;
            if (gsub == null || gsub.LookupList == null) return;

            int previousGlyphCount;
            do
            {
                previousGlyphCount = context.IncludedGlyphs.Count;

                foreach (var lookup in gsub.LookupList.Lookups)
                {
                    if (_handlers.TryGetValue(lookup.LookupType, out var handler))
                    {
                        handler.Discover(context, lookup);
                    }
                }

                // Continue as long as new glyphs are found (e.g., a ligature that triggers another substitution rule)
            } while (context.IncludedGlyphs.Count > previousGlyphCount);
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var oldGsub = context.OriginalFont.GsubTable;
            if (oldGsub == null) return;

            var newGsub = new GsubTable();

            // Clear internal mappings to prevent leakage between subsetting operations
            _oldToNewLookupIndex.Clear();
            _oldToNewFeatureIndex.Clear();

            // 1. Remap LookupList first (Features depend on these indices)
            newGsub.LookupList = RemapLookupListTable(context, oldGsub.LookupList);

            // 2. Remap FeatureList (Scripts depend on these indices)
            newGsub.FeatureList = RemapFeatureList(context, oldGsub.FeatureList);

            // 3. Remap ScriptList
            newGsub.ScriptList = RemapScriptList(context, oldGsub.ScriptList);

            context.SubsetFont.AddOrReplaceTable(newGsub);
        }

        private LookupListTable RemapLookupListTable(FontSubsettingContext context, LookupListTable oldList)
        {
            var newList = new LookupListTable();
            if (oldList == null) return newList;

            for (int i = 0; i < oldList.Lookups.Count; i++)
            {
                var oldLookup = oldList.Lookups[i];

                // Delegate the specific substitution logic to the registered handler
                if (_handlers.TryGetValue(oldLookup.LookupType, out var handler))
                {
                    var newLookup = handler.Rewrite(context, oldLookup);
                    if (newLookup != null)
                    {
                        newList.Lookups.Add(newLookup);
                        _oldToNewLookupIndex[i] = newList.Lookups.Count - 1;
                    }
                }
            }
            return newList;
        }

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

            // OpenType specification requires features to be sorted by tag
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

            // OpenType specification requires scripts to be sorted by tag
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