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
  01/09/2026         EPPlus Software AB           Enhanced validation for all GSUB types
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Common.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Features;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System;
using System.Collections.Generic;
using System.Linq;

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

            // 1. Validate version
            if (table.MajorVersion != 1)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"GSUB version {table.MajorVersion}.{table.MinorVersion} is not supported. Expected 1.x");
                return result; // Cannot continue if version is wrong
            }

            // 2. Validate ScriptListTable
            ValidateScriptList(table.ScriptList, result);

            // 3. Validate FeatureListTable
            int totalLookups = table.LookupList?.Lookups.Count ?? 0;
            ValidateFeatureList(table.FeatureList, result, totalLookups);

            // 4. Validate LookupListTable
            ValidateLookupList(table.LookupList, result, numGlyphs);

            return result;
        }

        #region ScriptList Validation

        private void ValidateScriptList(ScriptListTable scriptList, TableValidationResult result)
        {
            if (scriptList == null || scriptList.ScriptRecords == null)
            {
                result.AddMessage(FontValidationSeverity.Warning, "GSUB ScriptListTable is missing or null.");
                return;
            }

            if (scriptList.ScriptRecords.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning, "GSUB ScriptListTable has no scripts.");
                return;
            }

            for (int i = 0; i < scriptList.ScriptRecords.Count; i++)
            {
                var record = scriptList.ScriptRecords[i];

                // Validate the Tag
                if (record.ScriptTag == null || string.IsNullOrEmpty(record.ScriptTag.Value))
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"ScriptRecord at index {i} has a null or empty ScriptTag.");
                }
                else if (record.ScriptTag.Value.Length != 4)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"ScriptRecord index {i} has invalid tag length: '{record.ScriptTag.Value}' (must be 4 characters).");
                }

                // Check alphabetical sorting of tags (required by OpenType spec)
                if (i > 0 && record.ScriptTag != null && scriptList.ScriptRecords[i - 1].ScriptTag != null)
                {
                    if (string.CompareOrdinal(scriptList.ScriptRecords[i - 1].ScriptTag.Value, record.ScriptTag.Value) >= 0)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"ScriptRecords are not sorted alphabetically: '{scriptList.ScriptRecords[i - 1].ScriptTag.Value}' before '{record.ScriptTag.Value}'.");
                    }
                }

                // Validate the associated ScriptTable
                if (record.ScriptTable == null)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"ScriptRecord '{record.ScriptTag?.Value}' is missing its ScriptTable.");
                    continue;
                }

                ValidateScriptTable(record.ScriptTable, record.ScriptTag?.Value, result);
            }
        }

        private void ValidateScriptTable(ScriptTable scriptTable, string scriptTag, TableValidationResult result)
        {
            if (scriptTable.DefaultLangSys == null &&
                (scriptTable.LangSysRecords == null || scriptTable.LangSysRecords.Count == 0))
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Script '{scriptTag}' has no language systems (no DefaultLangSys and no LangSysRecords).");
                return;
            }

            // Validate DefaultLangSys
            if (scriptTable.DefaultLangSys != null)
            {
                ValidateLangSys(scriptTable.DefaultLangSys, $"{scriptTag} (default)", result);
            }

            // Validate LangSysRecords
            if (scriptTable.LangSysRecords != null)
            {
                for (int i = 0; i < scriptTable.LangSysRecords.Count; i++)
                {
                    var langRecord = scriptTable.LangSysRecords[i];

                    if (langRecord.LangSysTable != null)
                    {
                        ValidateLangSys(langRecord.LangSysTable,
                            $"{scriptTag}/{langRecord.LangSysTag:X8}", result);
                    }

                    // Check for sorted language tags
                    if (i > 0)
                    {
                        uint prevTag = scriptTable.LangSysRecords[i - 1].LangSysTag;
                        uint currTag = langRecord.LangSysTag;
                        if (prevTag >= currTag)
                        {
                            result.AddMessage(FontValidationSeverity.Error,
                                $"Script '{scriptTag}': LangSysRecords not sorted (0x{prevTag:X8} >= 0x{currTag:X8}).");
                        }
                    }
                }
            }
        }

        private void ValidateLangSys(LangSysTable langSys, string context, TableValidationResult result)
        {
            // LookupOrder must be 0 (reserved)
            if (langSys.LookupOrder != 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"{context}: LookupOrder is {langSys.LookupOrder} (should be 0 - reserved field).");
            }

            // Validate FeatureIndices
            if (langSys.FeatureIndices != null)
            {
                for (int i = 0; i < langSys.FeatureIndices.Length; i++)
                {
                    // Check for sorted indices (should be ascending)
                    if (i > 0 && langSys.FeatureIndices[i] <= langSys.FeatureIndices[i - 1])
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            $"{context}: FeatureIndices not sorted at index {i}.");
                    }
                }
            }
        }

        #endregion

        #region FeatureList Validation

        private void ValidateFeatureList(FeatureListTable featureList, TableValidationResult result, int totalLookups)
        {
            if (featureList == null || featureList.FeatureRecords == null)
            {
                result.AddMessage(FontValidationSeverity.Warning, "GSUB FeatureListTable is missing or null.");
                return;
            }

            for (int i = 0; i < featureList.FeatureRecords.Count; i++)
            {
                var record = featureList.FeatureRecords[i];

                // Validate Tag
                if (record.FeatureTag == null || string.IsNullOrEmpty(record.FeatureTag.Value))
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"FeatureRecord at index {i} has null or empty tag.");
                    continue;
                }

                if (record.FeatureTag.Value.Length != 4)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"FeatureRecord at index {i} has invalid tag length: '{record.FeatureTag.Value}'.");
                }

                // Validate Tag sorting (features should be sorted alphabetically)
                if (i > 0 && featureList.FeatureRecords[i - 1].FeatureTag != null)
                {
                    if (string.CompareOrdinal(featureList.FeatureRecords[i - 1].FeatureTag.Value,
                                             record.FeatureTag.Value) > 0)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"FeatureRecords are not sorted alphabetically: '{featureList.FeatureRecords[i - 1].FeatureTag.Value}' before '{record.FeatureTag.Value}'.");
                    }
                }

                // Validate FeatureTable
                if (record.FeatureTable == null)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Feature '{record.FeatureTag.Value}' has null FeatureTable.");
                    continue;
                }

                ValidateFeatureTable(record.FeatureTable, record.FeatureTag.Value, result, totalLookups);
            }
        }

        private void ValidateFeatureTable(FeatureTable featureTable, string featureTag,
            TableValidationResult result, int totalLookups)
        {
            if (featureTable.LookupListIndices == null || featureTable.LookupListIndices.Length == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Feature '{featureTag}' has no lookup indices (will have no effect).");
                return;
            }

            foreach (var lookupIdx in featureTable.LookupListIndices)
            {
                if (lookupIdx >= totalLookups)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Feature '{featureTag}' references out-of-bounds LookupIndex {lookupIdx} (max: {totalLookups - 1}).");
                }
            }
        }

        #endregion

        #region LookupList Validation

        private void ValidateLookupList(LookupListTable lookupList, TableValidationResult result, int numGlyphs)
        {
            if (lookupList == null || lookupList.Lookups == null)
            {
                result.AddMessage(FontValidationSeverity.Warning, "GSUB LookupListTable is missing or null.");
                return;
            }

            for (int i = 0; i < lookupList.Lookups.Count; i++)
            {
                var lookup = lookupList.Lookups[i];

                // Validate LookupType
                if (lookup.LookupType < 1 || lookup.LookupType > 8)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Lookup {i}: Invalid LookupType {lookup.LookupType} (must be 1-8).");
                    continue;
                }

                // Validate LookupFlags
                ValidateLookupFlags(lookup, i, result);

                // Validate SubTables
                if (lookup.SubTables == null || lookup.SubTables.Count == 0)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Lookup {i} (Type: {lookup.LookupType}) has no subtables (will have no effect).");
                    continue;
                }

                foreach (var subtable in lookup.SubTables)
                {
                    ValidateSubtable(subtable, lookup.LookupType, i, result, numGlyphs);
                }
            }
        }

        private void ValidateLookupFlags(LookupTable lookup, int lookupIdx, TableValidationResult result)
        {
            // Check reserved bits (bits 8-15 should be 0)
            if ((lookup.LookupFlag & 0xFF00) != 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Lookup {lookupIdx}: Reserved high bits in LookupFlag are set (0x{lookup.LookupFlag:X4}).");
            }

            // If UseMarkFilteringSet flag (0x0010) is set, MarkFilteringSet should exist
            if ((lookup.LookupFlag & 0x0010) != 0)
            {
                if (lookup.MarkFilteringSet == null)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Lookup {lookupIdx}: UseMarkFilteringSet flag is set but MarkFilteringSet is null.");
                }
            }
        }

        #endregion

        #region Subtable Validation

        private void ValidateSubtable(FontTableElement subtable, ushort lookupType, int lookupIdx,
    TableValidationResult result, int numGlyphs)
        {
            if (subtable == null)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx}: Subtable is null.");
                return;
            }

            if (subtable is SingleSubstSubTable single)
            {
                ValidateSingleSubst(single, lookupIdx, result, numGlyphs);
            }
            else if (subtable is LigatureSubstSubTable lig)
            {
                ValidateLigatureSubst(lig, lookupIdx, result, numGlyphs);
            }
            else if (subtable is ChainingContextualSubstFormat3 chain)
            {
                ValidateChainingContextualFormat3(chain, lookupIdx, result, numGlyphs);
            }
            else if (subtable is ExtensionSubstSubTable ext)
            {
                ValidateExtensionSubst(ext, lookupIdx, result, numGlyphs);
            }
            else
            {
                // Unknown or unimplemented subtable type
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Lookup {lookupIdx} (Type {lookupType}): Subtable type {subtable.GetType().Name} validation not implemented.");
            }
        }

        private void ValidateChainingContextualFormat3(ChainingContextualSubstFormat3 chain, int lookupIdx,
            TableValidationResult result, int numGlyphs)
        {
            // Validate Input coverages (required)
            if (chain.InputCoverages == null || chain.InputCoverages.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx}: ChainContext Format3 has no Input coverages (required).");
                return;
            }

            // Validate all coverage tables
            int covIdx = 0;
            if (chain.BacktrackCoverages != null)
            {
                foreach (var cov in chain.BacktrackCoverages)
                {
                    ValidateCoverage(cov, $"Lookup {lookupIdx} ChainContext Backtrack[{covIdx++}]", result, numGlyphs);
                }
            }

            covIdx = 0;
            foreach (var cov in chain.InputCoverages)
            {
                ValidateCoverage(cov, $"Lookup {lookupIdx} ChainContext Input[{covIdx++}]", result, numGlyphs);
            }

            covIdx = 0;
            if (chain.LookaheadCoverages != null)
            {
                foreach (var cov in chain.LookaheadCoverages)
                {
                    ValidateCoverage(cov, $"Lookup {lookupIdx} ChainContext Lookahead[{covIdx++}]", result, numGlyphs);
                }
            }

            // Validate SubstLookupRecords
            if (chain.SubstLookupRecords != null)
            {
                foreach (var record in chain.SubstLookupRecords)
                {
                    if (record.SequenceIndex >= chain.InputCoverages.Count)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Lookup {lookupIdx}: SubstLookupRecord SequenceIndex {record.SequenceIndex} exceeds Input count {chain.InputCoverages.Count}.");
                    }
                }
            }
        }

        #endregion

        #region Single Substitution Validation

        private void ValidateSingleSubst(SingleSubstSubTable single, int lookupIdx,
            TableValidationResult result, int numGlyphs)
        {
            ValidateCoverage(single.Coverage, $"Lookup {lookupIdx} (SingleSubst)", result, numGlyphs);

            if (single is SingleSubstSubTableFormat1 f1)
            {
                ValidateSingleSubstFormat1(f1, lookupIdx, result, numGlyphs);
            }
            else if (single is SingleSubstSubTableFormat2 f2)
            {
                ValidateSingleSubstFormat2(f2, single.Coverage, lookupIdx, result, numGlyphs);
            }
            else
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx}: Unknown SingleSubst format.");
            }
        }

        private void ValidateSingleSubstFormat1(SingleSubstSubTableFormat1 f1, int lookupIdx,
            TableValidationResult result, int numGlyphs)
        {
            // Validate that delta doesn't produce out-of-range glyphs
            var coveredGlyphs = f1.Coverage?.GetCoveredGlyphs() ?? new ushort[0];

            foreach (var gid in coveredGlyphs)
            {
                // DeltaGlyphID is signed, so cast for proper arithmetic
                int resultGid = gid + (short)f1.DeltaGlyphID;

                if (resultGid < 0)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Lookup {lookupIdx} Format1: Delta {(short)f1.DeltaGlyphID} produces negative glyph ID for glyph {gid} → {resultGid}.");
                }
                else if (resultGid >= numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Lookup {lookupIdx} Format1: Delta {(short)f1.DeltaGlyphID} produces out-of-range glyph ID for glyph {gid} → {resultGid} (max: {numGlyphs - 1}).");
                }
            }
        }

        private void ValidateSingleSubstFormat2(SingleSubstSubTableFormat2 f2, CoverageTable coverage,
            int lookupIdx, TableValidationResult result, int numGlyphs)
        {
            if (f2.SubstituteGlyphIDs == null || f2.SubstituteGlyphIDs.Length == 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx} Format2: SubstituteGlyphIDs is null or empty.");
                return;
            }

            // Validate each substitute glyph ID
            foreach (var gid in f2.SubstituteGlyphIDs)
            {
                if (gid >= numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Lookup {lookupIdx} Format2: Substitute glyph ID {gid} is out of range (max: {numGlyphs - 1}).");
                }
            }

            // Validate count matches coverage
            int coverageCount = GetCoverageGlyphCount(coverage);
            if (f2.SubstituteGlyphIDs.Length != coverageCount)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx} Format2: SubstituteGlyphIDs count ({f2.SubstituteGlyphIDs.Length}) does not match Coverage count ({coverageCount}).");
            }
        }

        #endregion

        #region Ligature Substitution Validation

        private void ValidateLigatureSubst(LigatureSubstSubTable lig, int lookupIdx,
            TableValidationResult result, int numGlyphs)
        {
            ValidateCoverage(lig.Coverage, $"Lookup {lookupIdx} (LigatureSubst)", result, numGlyphs);

            if (lig.LigatureSets == null || lig.LigatureSets.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Lookup {lookupIdx}: LigatureSubst has no ligature sets.");
                return;
            }

            foreach (var kvp in lig.LigatureSets)
            {
                ushort firstGlyph = kvp.Key;
                var ligSet = kvp.Value;

                if (ligSet == null || ligSet.Ligatures == null || ligSet.Ligatures.Count == 0)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Lookup {lookupIdx}: LigatureSet for glyph {firstGlyph} is empty.");
                    continue;
                }

                foreach (var ligature in ligSet.Ligatures)
                {
                    ValidateLigatureTable(ligature, firstGlyph, lookupIdx, result, numGlyphs);
                }
            }
        }

        private void ValidateLigatureTable(LigatureTable ligature, ushort firstGlyph, int lookupIdx,
            TableValidationResult result, int numGlyphs)
        {
            // Validate output glyph
            if (ligature.LigatureGlyph >= numGlyphs)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx}: Ligature output glyph {ligature.LigatureGlyph} is out of range (max: {numGlyphs - 1}).");
            }

            // Validate component count
            if (ligature.Components == null || ligature.Components.Length == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Lookup {lookupIdx}: Ligature {firstGlyph} → {ligature.LigatureGlyph} has no components (should have at least 1).");
                return;
            }

            // Validate each component
            foreach (var compGid in ligature.Components)
            {
                if (compGid >= numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Lookup {lookupIdx}: Ligature component glyph {compGid} is out of range (max: {numGlyphs - 1}).");
                }
            }

            // Check for potential circular references
            if (ligature.LigatureGlyph >= 400)
            {
                foreach (var compGid in ligature.Components)
                {
                    if (compGid >= 400)
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            $"Lookup {lookupIdx}: Ligature {ligature.LigatureGlyph} uses component {compGid} which may be another ligature (potential circular dependency).");
                    }
                }
            }

            // Check for self-reference
            if (firstGlyph == ligature.LigatureGlyph || ligature.Components.Contains(ligature.LigatureGlyph))
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx}: Ligature {ligature.LigatureGlyph} references itself (circular dependency).");
            }
        }

        #endregion

        #region Extension Substitution Validation

        private void ValidateExtensionSubst(ExtensionSubstSubTable ext, int lookupIdx,
     TableValidationResult result, int numGlyphs)
        {
            // Extension cannot wrap another extension
            if (ext.ExtensionLookupType == 7)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx}: Extension substitution cannot reference another extension (Type 7).");
                return;
            }

            // Validate extension lookup type range
            if (ext.ExtensionLookupType < 1 || ext.ExtensionLookupType > 8)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx}: Invalid ExtensionLookupType {ext.ExtensionLookupType} (must be 1-8, not 7).");
                return;
            }

            // Validate wrapped subtable exists
            if (ext.ExtendedSubTable == null)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Lookup {lookupIdx}: Extension has no wrapped subtable.");
                return;
            }

            // Recursively validate the wrapped subtable
            ValidateSubtable(ext.ExtendedSubTable, ext.ExtensionLookupType, lookupIdx, result, numGlyphs);
        }

        #endregion

        #region Coverage Validation

        private void ValidateCoverage(CoverageTable coverage, string context,
            TableValidationResult result, int numGlyphs)
        {
            if (coverage == null)
            {
                result.AddMessage(FontValidationSeverity.Error, $"{context}: CoverageTable is missing.");
                return;
            }

            if (coverage is CoverageTableFormat1 f1)
            {
                ValidateCoverageFormat1(f1, context, result, numGlyphs);
            }
            else if (coverage is CoverageTableFormat2 f2)
            {
                ValidateCoverageFormat2(f2, context, result, numGlyphs);
            }
            else
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"{context}: Unknown CoverageTable format.");
            }
        }

        private void ValidateCoverageFormat1(CoverageTableFormat1 f1, string context,
            TableValidationResult result, int numGlyphs)
        {
            if (f1.GlyphArray == null || f1.GlyphArray.Length == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"{context}: Coverage Format1 has empty GlyphArray.");
                return;
            }

            for (int i = 0; i < f1.GlyphArray.Length; i++)
            {
                var gid = f1.GlyphArray[i];

                // Check range
                if (gid >= numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"{context}: Coverage glyph ID {gid} at index {i} is out of range (max: {numGlyphs - 1}).");
                }

                // Check strict ascending order (required by spec)
                if (i > 0 && gid <= f1.GlyphArray[i - 1])
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"{context}: Coverage GlyphArray not strictly ascending at index {i} ({f1.GlyphArray[i - 1]} → {gid}).");
                }
            }
        }

        private void ValidateCoverageFormat2(CoverageTableFormat2 f2, string context,
            TableValidationResult result, int numGlyphs)
        {
            if (f2.RangeRecords == null || f2.RangeRecords.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"{context}: Coverage Format2 has no RangeRecords.");
                return;
            }

            for (int i = 0; i < f2.RangeRecords.Count; i++)
            {
                var range = f2.RangeRecords[i];

                // Check range validity
                if (range.StartGlyphID > range.EndGlyphID)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"{context}: RangeRecord {i} has StartGlyphID ({range.StartGlyphID}) > EndGlyphID ({range.EndGlyphID}).");
                }

                // Check glyph IDs are in range
                if (range.StartGlyphID >= numGlyphs || range.EndGlyphID >= numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"{context}: RangeRecord {i} contains out-of-range glyph IDs ({range.StartGlyphID}-{range.EndGlyphID}, max: {numGlyphs - 1}).");
                }

                // Check for overlapping or unsorted ranges
                if (i > 0)
                {
                    var prevRange = f2.RangeRecords[i - 1];
                    if (range.StartGlyphID <= prevRange.EndGlyphID)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"{context}: RangeRecords overlap or not sorted at index {i} (prev: {prevRange.StartGlyphID}-{prevRange.EndGlyphID}, curr: {range.StartGlyphID}-{range.EndGlyphID}).");
                    }
                }
            }
        }

        private int GetCoverageGlyphCount(CoverageTable coverage)
        {
            if (coverage is CoverageTableFormat1 f1)
            {
                return f1.GlyphArray?.Length ?? 0;
            }

            if (coverage is CoverageTableFormat2 f2)
            {
                int count = 0;
                if (f2.RangeRecords != null)
                {
                    foreach (var r in f2.RangeRecords)
                    {
                        count += (r.EndGlyphID - r.StartGlyphID + 1);
                    }
                }
                return count;
            }

            return 0;
        }

        #endregion
    }
}