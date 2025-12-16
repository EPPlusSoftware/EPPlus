using EPPlus.Fonts.OpenType.FontValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    internal class GsubTableValidator : TableValidatorBase<GsubTable>
    {
        public override Type TableType => typeof(GsubTable);
        public override string TableName => TableNames.Gsub;

        public override TableValidationResult Validate(GsubTable table, FontValidationContext context)
        {
            var result = new TableValidationResult { TableName = TableName, LogLevel = base.LogLevel };
            var numGlyphs = context.Font.MaxpTable?.numGlyphs ?? 0;

            if (table == null) return result;

            // 1. Validate ScriptListTable
            ValidateScriptList(table.ScriptList, result);

            // 2. Validate FeatureListTable
            int totalLookups = table.LookupList?.Lookups.Count ?? 0;
            ValidateFeatureList(table.FeatureList, result, totalLookups);

            // 3. Validate LookupListTable
            ValidateLookupList(table.LookupList, result, numGlyphs);

            return result;
        }

        private void ValidateScriptList(ScriptListTable scriptList, TableValidationResult result)
        {
            if (scriptList == null || scriptList.ScriptRecords == null)
            {
                result.AddMessage(FontValidationSeverity.Warning, "GSUB ScriptListTable is missing or null.");
                return;
            }

            for (int i = 0; i < scriptList.ScriptRecords.Count; i++)
            {
                var record = scriptList.ScriptRecords[i];

                // Validate the Tag
                if (record.ScriptTag == null || string.IsNullOrEmpty(record.ScriptTag.Value))
                {
                    result.AddMessage(FontValidationSeverity.Error, $"ScriptRecord at index {i} has a null or empty ScriptTag.");
                }
                else if (record.ScriptTag.Value.Length != 4)
                {
                    result.AddMessage(FontValidationSeverity.Error, $"ScriptRecord index {i} has invalid tag length: '{record.ScriptTag.Value}'.");
                }

                // Check alphabetical sorting of tags (required by OpenType spec)
                if (i > 0 && record.ScriptTag != null && scriptList.ScriptRecords[i - 1].ScriptTag != null)
                {
                    if (string.CompareOrdinal(scriptList.ScriptRecords[i - 1].ScriptTag.Value, record.ScriptTag.Value) >= 0)
                    {
                        result.AddMessage(FontValidationSeverity.Error, $"ScriptRecords are not sorted alphabetically: '{scriptList.ScriptRecords[i - 1].ScriptTag.Value}' before '{record.ScriptTag.Value}'.");
                    }
                }

                // Validate the associated ScriptTable
                if (record.ScriptTable == null)
                {
                    result.AddMessage(FontValidationSeverity.Error, $"ScriptRecord '{record.ScriptTag?.Value}' is missing its ScriptTable.");
                    continue;
                }

                if (record.ScriptTable.DefaultLangSys == null && (record.ScriptTable.LangSysRecords == null || record.ScriptTable.LangSysRecords.Count == 0))
                {
                    result.AddMessage(FontValidationSeverity.Warning, $"Script '{record.ScriptTag?.Value}' has no language systems.");
                }
            }
        }

        private void ValidateFeatureList(FeatureListTable featureList, TableValidationResult result, int totalLookups)
        {
            if (featureList == null || featureList.FeatureRecords == null) return;

            for (int i = 0; i < featureList.FeatureRecords.Count; i++)
            {
                var record = featureList.FeatureRecords[i];

                // Validate Tag sorting for Features
                if (i > 0 && record.FeatureTag != null && featureList.FeatureRecords[i - 1].FeatureTag != null)
                {
                    if (string.CompareOrdinal(featureList.FeatureRecords[i - 1].FeatureTag.Value, record.FeatureTag.Value) > 0)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"FeatureRecords are not sorted alphabetically: '{featureList.FeatureRecords[i - 1].FeatureTag.Value}' before '{record.FeatureTag.Value}'.");
                    }
                }

                if (record.FeatureTable == null) continue;

                foreach (var lookupIdx in record.FeatureTable.LookupListIndices)
                {
                    if (lookupIdx >= totalLookups)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Feature '{record.FeatureTag?.Value}' references out-of-bounds LookupIndex {lookupIdx}.");
                    }
                }
            }
        }

        private void ValidateLookupList(LookupListTable lookupList, TableValidationResult result, int numGlyphs)
        {
            if (lookupList == null || lookupList.Lookups == null) return;

            for (int i = 0; i < lookupList.Lookups.Count; i++)
            {
                var lookup = lookupList.Lookups[i];
                if (lookup.SubTables == null || lookup.SubTables.Count == 0)
                {
                    result.AddMessage(FontValidationSeverity.Warning, "Lookup index " + i + " (Type: " + lookup.LookupType + ") has no subtables. This lookup will have no effect.");
                    continue;
                }

                foreach (var subtable in lookup.SubTables)
                {
                    ValidateSubtable(subtable, i, result, numGlyphs);
                }
            }
        }

        private void ValidateSubtable(FontTableElement subtable, int lookupIdx, TableValidationResult result, int numGlyphs)
        {
            if (subtable is SingleSubstSubTable single)
            {
                ValidateCoverage(single.Coverage, $"Lookup {lookupIdx} (SingleSubst)", result, numGlyphs);

                if (subtable is SingleSubstSubTableFormat2 f2)
                {
                    if (f2.SubstituteGlyphIDs != null)
                    {
                        foreach (var gid in f2.SubstituteGlyphIDs)
                        {
                            if (gid >= numGlyphs)
                                result.AddMessage(FontValidationSeverity.Error, $"Lookup {lookupIdx}: Substitute glyph ID {gid} is out of range.");
                        }

                        int coverageCount = GetCoverageGlyphCount(single.Coverage);
                        if (f2.SubstituteGlyphIDs.Length != coverageCount)
                        {
                            result.AddMessage(FontValidationSeverity.Error,
                                $"Lookup {lookupIdx}: SubstituteGlyphIDs count ({f2.SubstituteGlyphIDs.Length}) does not match Coverage count ({coverageCount}).");
                        }
                    }
                }
            }
            else if (subtable is LigatureSubstSubTable lig)
            {
                ValidateCoverage(lig.Coverage, $"Lookup {lookupIdx} (LigatureSubst)", result, numGlyphs);

                if (lig.LigatureSets == null) return;

                foreach (var kvp in lig.LigatureSets)
                {
                    if (kvp.Value?.Ligatures == null) continue;

                    foreach (var l in kvp.Value.Ligatures)
                    {
                        if (l.LigatureGlyph >= numGlyphs)
                            result.AddMessage(FontValidationSeverity.Error, $"Lookup {lookupIdx}: Ligature output ID {l.LigatureGlyph} is out of range.");

                        if (l.Components != null)
                        {
                            foreach (var compGid in l.Components)
                            {
                                if (compGid >= numGlyphs)
                                    result.AddMessage(FontValidationSeverity.Error, $"Lookup {lookupIdx}: Ligature component ID {compGid} is out of range.");
                            }
                        }
                    }
                }
            }
        }

        private int GetCoverageGlyphCount(CoverageTable coverage)
        {
            if (coverage is CoverageTableFormat1 f1) return f1.GlyphArray?.Length ?? 0;
            if (coverage is CoverageTableFormat2 f2)
            {
                int count = 0;
                if (f2.RangeRecords != null)
                {
                    foreach (var r in f2.RangeRecords)
                        count += (r.EndGlyphID - r.StartGlyphID + 1);
                }
                return count;
            }
            return 0;
        }

        private void ValidateCoverage(CoverageTable coverage, string context, TableValidationResult result, int numGlyphs)
        {
            if (coverage == null)
            {
                result.AddMessage(FontValidationSeverity.Error, $"{context}: CoverageTable is missing.");
                return;
            }

            if (coverage is CoverageTableFormat1 f1 && f1.GlyphArray != null)
            {
                for (int i = 0; i < f1.GlyphArray.Length; i++)
                {
                    var gid = f1.GlyphArray[i];
                    if (gid >= numGlyphs)
                        result.AddMessage(FontValidationSeverity.Error, $"{context}: Coverage ID {gid} is out of range.");

                    if (i > 0 && gid <= f1.GlyphArray[i - 1])
                        result.AddMessage(FontValidationSeverity.Error, $"{context}: Coverage GlyphArray is not strictly ascending at index {i}.");
                }
            }
            else if (coverage is CoverageTableFormat2 f2 && f2.RangeRecords != null)
            {
                for (int i = 0; i < f2.RangeRecords.Count; i++)
                {
                    var range = f2.RangeRecords[i];
                    if (range.StartGlyphID >= numGlyphs || range.EndGlyphID >= numGlyphs)
                        result.AddMessage(FontValidationSeverity.Error, $"{context}: RangeRecord {i} contains out-of-range IDs.");

                    if (range.StartGlyphID > range.EndGlyphID)
                        result.AddMessage(FontValidationSeverity.Error, $"{context}: RangeRecord {i} has StartGlyphID > EndGlyphID.");

                    if (i > 0 && range.StartGlyphID <= f2.RangeRecords[i - 1].EndGlyphID)
                        result.AddMessage(FontValidationSeverity.Error, $"{context}: RangeRecords are overlapping or unsorted at index {i}.");
                }
            }
        }
    }
}