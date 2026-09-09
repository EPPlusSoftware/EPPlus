/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Ligature feature-tag plumbing (WP3/WP4)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Features;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.TextShaping.Ligatures;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping.Ligatures
{
    /// <summary>
    /// Covers two previous defects in LigatureProcessor:
    ///
    ///   1. The feature tag "liga" used to be hardcoded in the constructor, so a ligature tagged
    ///      "dlig", "clig" or "rlig" was never even loaded, regardless of what
    ///      ShapingOptions.GsubFeatures asked for.
    ///   2. "if (lookup.LookupType != 4) continue" used to skip lookup type 7 (Extension
    ///      Substitution) entirely. Unlike GPOS, the GSUB loader does NOT unwrap extension
    ///      lookups - the wrapper survives with LookupType still 7 - so an extension-wrapped
    ///      ligature was silently dropped.
    ///
    /// Both tests use a synthetic GsubTable so they do not depend on any specific font file.
    /// </summary>
    [TestClass]
    public class LigatureFeatureTagTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const ushort GlyphA = 5;
        private const ushort GlyphB = 6;
        private const ushort LigatureGlyph = 7;

        /// <summary>
        /// Previously the constructor only ever looked for the "liga" tag, regardless of what was
        /// requested at call time, so a ligature tagged "dlig" was invisible to it no matter what.
        /// Now that LigatureProcessor's constructor scans every tag and ApplyLigaturesInPlace
        /// takes a featureTags parameter (matching the pattern already used by
        /// SingleSubstitutionProcessor.ApplySubstitutions and
        /// ChainingContextualProcessor.ApplyContextualSubstitutions), this is a regression test.
        /// </summary>
        [TestMethod]
        public void Ligature_TaggedDlig_IsAppliedWhenRequested()
        {
            var font = TestFolderEngine.LoadFont("BIZUDGothic", FontSubFamily.Regular, ignoreCache: true);
            Assert.IsNotNull(font, "BIZUDGothic must be present in the test font folder");

            font.AddOrReplaceTable(BuildDligOnlyGsubTable());

            var processor = new LigatureProcessor(font);

            var glyphs = new List<ShapedGlyph>
            {
                new ShapedGlyph { GlyphId = GlyphA },
                new ShapedGlyph { GlyphId = GlyphB }
            };

            processor.ApplyLigaturesInPlace(glyphs, new List<string> { "dlig" }, "latn", null);

            Assert.AreEqual(
                1,
                glyphs.Count,
                "A+B must be merged into the ligature when \"dlig\" is an active feature tag, "
                + "even though the ligature is not tagged \"liga\"");
            Assert.AreEqual(LigatureGlyph, glyphs[0].GlyphId);
        }

        /// <summary>
        /// Now a regression test. The "liga" lookup here is Extension Substitution (Type 7) wrapping a
        /// LigatureSubstSubTable, which is exactly how EBGaramond and other real fonts in the test
        /// suite store some of their ligature data. "if (lookup.LookupType != 4) continue" skips
        /// it outright, so the A+B ligature never fires despite "liga" being requested and present.
        /// </summary>
        [TestMethod]
        public void Ligature_ExtensionWrapped_IsApplied()
        {
            var font = TestFolderEngine.LoadFont("BIZUDGothic", FontSubFamily.Regular, ignoreCache: true);
            Assert.IsNotNull(font, "BIZUDGothic must be present in the test font folder");

            font.AddOrReplaceTable(BuildExtensionWrappedLigaGsubTable());

            var processor = new LigatureProcessor(font);

            var glyphs = new List<ShapedGlyph>
            {
                new ShapedGlyph { GlyphId = GlyphA },
                new ShapedGlyph { GlyphId = GlyphB }
            };

            // Uses the same signature as the "dlig" test above: this defect is independent of
            // the feature-tag plumbing fix and is reachable through the "liga" tag alone.
            processor.ApplyLigaturesInPlace(glyphs, new List<string> { "liga" }, "latn", null);

            Assert.AreEqual(
                1,
                glyphs.Count,
                "A+B must be merged even though the \"liga\" lookup is Extension Substitution "
                + "(Type 7) wrapping a LigatureSubstSubTable, not a direct Type 4 lookup");
            Assert.AreEqual(LigatureGlyph, glyphs[0].GlyphId);
        }

        /// <summary>
        /// One "dlig" FeatureRecord -> a direct Type 4 LigatureSubstSubTable merging A+B.
        /// No "liga" tag anywhere in this table.
        /// </summary>
        private static GsubTable BuildDligOnlyGsubTable()
        {
            var ligatureSubtable = BuildLigatureSubtable();

            var lookup = new LookupTable
            {
                LookupType = 4,
                SubTables = new List<Tables.FontTableElement> { ligatureSubtable }
            };

            var featureList = new FeatureListTable
            {
                FeatureRecords = new List<FeatureRecord>
                {
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("dlig"),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 0 } }
                    }
                }
            };

            // No ScriptList: GetActiveIndices falls back to "no filter", so this test exercises
            // only the feature-tag defect, not script filtering.
            return new GsubTable
            {
                FeatureList = featureList,
                LookupList = new LookupListTable { Lookups = new List<LookupTable> { lookup } }
            };
        }

        /// <summary>
        /// One "liga" FeatureRecord -> a Type 7 (Extension) lookup wrapping the same Type 4
        /// LigatureSubstSubTable, mirroring how real fonts store extension-wrapped GSUB data
        /// (GSUB keeps the wrapper; unlike GPOS, the loader does not flatten it).
        /// </summary>
        private static GsubTable BuildExtensionWrappedLigaGsubTable()
        {
            var ligatureSubtable = BuildLigatureSubtable();

            var extensionWrapper = new ExtensionSubstSubTable
            {
                ExtensionLookupType = 4,
                ExtendedSubTable = ligatureSubtable
            };

            var lookup = new LookupTable
            {
                LookupType = 7,
                SubTables = new List<Tables.FontTableElement> { extensionWrapper }
            };

            var featureList = new FeatureListTable
            {
                FeatureRecords = new List<FeatureRecord>
                {
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("liga"),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 0 } }
                    }
                }
            };

            return new GsubTable
            {
                FeatureList = featureList,
                LookupList = new LookupListTable { Lookups = new List<LookupTable> { lookup } }
            };
        }

        private static LigatureSubstSubTable BuildLigatureSubtable()
        {
            return new LigatureSubstSubTable
            {
                Coverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { GlyphA } },
                LigatureSets = new Dictionary<ushort, LigatureSetTable>
                {
                    [GlyphA] = new LigatureSetTable
                    {
                        Ligatures = new List<LigatureTable>
                        {
                            new LigatureTable { LigatureGlyph = LigatureGlyph, Components = new ushort[] { GlyphB } }
                        }
                    }
                }
            };
        }
    }
}