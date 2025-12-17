using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Handlers;
using System;
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

                // Fortsätt så länge vi hittar nya glyfer (t.ex. en ligatur som i sin tur kan vara del av en annan regel)
            } while (context.IncludedGlyphs.Count > previousGlyphCount);
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var oldGsub = context.OriginalFont.GsubTable;
            if (oldGsub == null) return;

            var newGsub = new GsubTable();

            // Rensa interna mappar för att undvika läckage mellan körningar
            _oldToNewLookupIndex.Clear();
            _oldToNewFeatureIndex.Clear();

            // 1. Mappa om LookupList först (eftersom Features pekar på dessa)
            // Vi skickar med context för att komma åt Glyph-mappningen (Old-to-New GID)
            newGsub.LookupList = RemapLookupListTable(context, oldGsub.LookupList);

            // 2. Mappa om FeatureList (eftersom Scripts pekar på dessa)
            newGsub.FeatureList = RemapFeatureList(context, oldGsub.FeatureList);

            // 3. Mappa om ScriptList
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

                // Här är hjärtat i vår pipeline:
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

        private LookupTable RewriteLookupTable(FontSubsettingContext context, LookupTable oldLookup)
        {
            // Skapa en ny instans av LookupTable för vår subset-font
            LookupTable newLookup = new LookupTable();
            newLookup.LookupType = oldLookup.LookupType;
            newLookup.LookupFlag = oldLookup.LookupFlag;
            newLookup.SubTables = new List<FontTableElement>();

            foreach (FontTableElement subtableElement in oldLookup.SubTables)
            {
                if((subtableElement is not SingleSubstSubTable) && (subtableElement is not LigatureSubstSubTable))
                {
                    int i2 = 0;
                }
                // 1. Hantera Single Substitution (Typ 1)
                SingleSubstSubTable singleSub = subtableElement as SingleSubstSubTable;
                if (singleSub != null)
                {
                    SingleSubstSubTable rewrittenSingle = singleSub.Rewrite(context);
                    if (rewrittenSingle != null)
                    {
                        newLookup.SubTables.Add(rewrittenSingle);
                    }
                    continue;
                }

                // 2. Hantera Ligature Substitution (Typ 4)
                LigatureSubstSubTable ligatureSub = subtableElement as LigatureSubstSubTable;
                if (ligatureSub != null)
                {
                    LigatureSubstSubTable rewrittenLig = ligatureSub.Rewrite(context);
                    if (rewrittenLig != null)
                    {
                        newLookup.SubTables.Add(rewrittenLig);
                    }
                    continue;
                }

                // Här kan du i framtiden lägga till fler typer, t.ex. MultipleSubst (Typ 2)
            }

            // Om inga subtabeller överlevde filtreringen (t.ex. om inga av tecknen i 
            // tabellen finns i vårt subset), returnerar vi null så att Lookupen kan rensas bort helt.
            return newLookup.SubTables.Count > 0 ? newLookup : null;
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