/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Script-aware feature resolution
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts
{
    /// <summary>
    /// Resolves which entries in a FeatureList a given script/language may use, per the
    /// ScriptList/LangSys structure in the OpenType spec.
    ///
    /// Every GSUB/GPOS processor used to iterate FeatureList.FeatureRecords directly and
    /// applied every record whose tag matched, regardless of script. That let a lookup reachable
    /// only through one script's LangSys (e.g. an Arabic-only "kern" adjustment) apply to text in
    /// a completely different script. Real-world example: a SinglePos reachable only via 'arab'
    /// gave the Latin 'period' glyph an unwanted YPlacement in one font's GPOS table.
    /// </summary>
    internal static class ScriptFeatureResolver
    {
        private const string DefaultScriptTag = "DFLT";

        /// <summary>
        /// Returns the set of FeatureList indices reachable from the given script and language.
        ///
        /// Returns null - meaning "no filter, keep every FeatureRecord" - when scriptList is null
        /// or the requested script is not null but neither the exact script nor 'DFLT' is present.
        /// Returning null instead of an empty set is deliberate: it preserves the previous behavior
        /// (apply every matching-tag record) for callers that have no way to resolve a script,
        /// rather than silently discarding all features. Every current caller passes "latn" from
        /// ShapingOptions.Default, so this fallback should not be reachable in practice today.
        /// </summary>
        /// <param name="scriptList">The font's ScriptList table, or null if unavailable.</param>
        /// <param name="script">
        /// OpenType script tag (e.g. "latn"). Case-insensitive; padded/truncated to 4 characters
        /// as OpenType tags require. Null means "use the font's default script".
        /// </param>
        /// <param name="language">
        /// OpenType language-system tag (e.g. "SWE "), or null for the script's default LangSys.
        /// </param>
        public static HashSet<int> GetActiveFeatureIndices(ScriptListTable scriptList, string script, string language)
        {
            if (scriptList == null)
            {
                return null;
            }

            var scriptTable = FindScriptTable(scriptList, script) ?? FindScriptTable(scriptList, DefaultScriptTag);
            if (scriptTable == null)
            {
                // Requested script (and DFLT) both absent from this font. No safe default exists,
                // so fall back to unfiltered rather than silently dropping every feature.
                return null;
            }

            var langSys = FindLangSys(scriptTable, language) ?? scriptTable.DefaultLangSys;
            if (langSys == null)
            {
                return new HashSet<int>();
            }

            var indices = new HashSet<int>();
            if (langSys.RequiredFeatureIndex != 0xFFFF)
            {
                indices.Add(langSys.RequiredFeatureIndex);
            }

            if (langSys.FeatureIndices != null)
            {
                foreach (var index in langSys.FeatureIndices)
                {
                    indices.Add(index);
                }
            }

            return indices;
        }

        private static ScriptTable FindScriptTable(ScriptListTable scriptList, string scriptTag)
        {
            if (string.IsNullOrEmpty(scriptTag))
            {
                return null;
            }

            foreach (var record in scriptList.ScriptRecords)
            {
                if (TagsMatch(record.ScriptTag?.Value, scriptTag))
                {
                    return record.ScriptTable;
                }
            }

            return null;
        }

        private static LangSysTable FindLangSys(ScriptTable scriptTable, string languageTag)
        {
            if (string.IsNullOrEmpty(languageTag) || scriptTable.LangSysRecords == null)
            {
                return null;
            }

            foreach (var record in scriptTable.LangSysRecords)
            {
                if (TagsMatch(LangSysTagToString(record.LangSysTag), languageTag))
                {
                    return record.LangSysTable;
                }
            }

            return null;
        }

        /// <summary>
        /// OpenType tags are compared case-insensitively here despite the spec's binary
        /// comparison, because callers pass human-typed strings (ShapingOptions.Script/.Language)
        /// rather than tags read verbatim from a font file.
        /// </summary>
        private static bool TagsMatch(string a, string b)
        {
            return a != null && b != null
                && string.Equals(a.TrimEnd(), b.TrimEnd(), StringComparison.OrdinalIgnoreCase);
        }

        private static string LangSysTagToString(uint tag)
        {
            var bytes = new[]
            {
                (char)((tag >> 24) & 0xFF),
                (char)((tag >> 16) & 0xFF),
                (char)((tag >> 8) & 0xFF),
                (char)(tag & 0xFF)
            };
            return new string(bytes);
        }
    }
}