/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Script-aware feature lookup
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Features;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.TextShaping.Contextual;
using EPPlus.Fonts.OpenType.TextShaping.Ligatures;
using EPPlus.Fonts.OpenType.TextShaping.Substitutions;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    /// <summary>
    /// Same synthetic-GsubTable technique as ScriptAwareFeatureLookupTests and
    /// ScriptAwareKerningAndMarkTests, applied to the three remaining GSUB processors:
    /// LigatureProcessor, ChainingContextualProcessor and SingleSubstitutionProcessor.
    ///
    /// Each test builds two FeatureRecords sharing a tag ("liga" / "smcp") but reachable through
    /// different scripts, with the 'arab'-only record placed last so it would win any tag-only
    /// (script-unaware) map regardless of which script is actually active.
    /// </summary>
    [TestClass]
    public class ScriptAwareGsubTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const ushort GlyphA = 5;
        private const ushort GlyphB = 6;
        private const ushort LigatureGlyph = 7;
        private const ushort SubstituteGlyph = 8;

        private static OpenTypeFont LoadHostFont()
        {
            var font = TestFolderEngine.LoadFont("BIZUDGothic", FontSubFamily.Regular, ignoreCache: true);
            Assert.IsNotNull(font, "BIZUDGothic must be present in the test font folder");
            return font;
        }

        [TestMethod]
        public void Ligatures_ForLatinScript_DoNotPickUpArabicOnlyLigature()
        {
            var font = LoadHostFont();
            font.AddOrReplaceTable(BuildLigatureGsubTable());

            var processor = new LigatureProcessor(font);

            var glyphs = new List<ShapedGlyph>
            {
                new ShapedGlyph { GlyphId = GlyphA },
                new ShapedGlyph { GlyphId = GlyphB }
            };

            processor.ApplyLigaturesInPlace(glyphs, new List<string> { "liga" }, "latn", null);

            Assert.AreEqual(
                2,
                glyphs.Count,
                "A+B must not be merged into the ligature that is only reachable through the "
                + "'arab' script's LangSys");
        }

        [TestMethod]
        public void ChainingContextual_ForLatinScript_DoesNotPickUpArabicOnlyRule()
        {
            var font = LoadHostFont();
            font.AddOrReplaceTable(BuildChainingContextualGsubTable());

            var processor = new ChainingContextualProcessor(
                font,
                new SingleSubstitutionProcessor(font),
                new LigatureProcessor(font));

            var glyphs = new List<ShapedGlyph>
            {
                new ShapedGlyph { GlyphId = GlyphA }
            };

            var result = processor.ApplyContextualSubstitutions(glyphs, "liga", "latn", null);

            Assert.AreEqual(
                GlyphA,
                result[0].GlyphId,
                "the glyph must not be substituted by the contextual rule that is only reachable "
                + "through the 'arab' script's LangSys");
        }

        [TestMethod]
        public void SingleSubstitution_ForLatinScript_DoesNotPickUpArabicOnlySubstitution()
        {
            var font = LoadHostFont();
            font.AddOrReplaceTable(BuildSingleSubstGsubTable());

            var processor = new SingleSubstitutionProcessor(font);

            var glyphs = new List<ShapedGlyph>
            {
                new ShapedGlyph { GlyphId = GlyphA, BaseAdvance = 500, XAdvance = 500 }
            };

            var result = processor.ApplySubstitutions(
                glyphs, new List<string> { "smcp" }, "latn", null);

            Assert.AreEqual(
                GlyphA,
                result[0].GlyphId,
                "the glyph must not be substituted by the 'smcp' record that is only reachable "
                + "through the 'arab' script's LangSys");
        }

        /// <summary>
        /// Two "liga" FeatureRecords: index 0 -> 'latn', a lookup with no subtables (matches
        /// nothing). Index 1 -> 'arab', a LigatureSubstSubTable merging GlyphA+GlyphB into
        /// LigatureGlyph. Index 1 is last, so it wins a tag-only map regardless of active script.
        /// </summary>
        private static GsubTable BuildLigatureGsubTable()
        {
            var arabLigature = new LigatureSubstSubTable
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

            var latinLookup = new LookupTable { LookupType = 4, SubTables = new List<FontTableElement>() };
            var arabLookup = new LookupTable
            {
                LookupType = 4,
                SubTables = new List<FontTableElement> { arabLigature }
            };

            return new GsubTable
            {
                ScriptList = BuildTwoScriptList(),
                FeatureList = BuildTwoFeatureList("liga"),
                LookupList = new LookupListTable { Lookups = new List<LookupTable> { latinLookup, arabLookup } }
            };
        }

        /// <summary>
        /// Two "liga" FeatureRecords (ChainingContextualProcessor is hardcoded to look for
        /// "liga", per LigatureProcessor's own hardcoding pre-WP3/WP4). Index 0 -> 'latn', a Type
        /// 6 lookup with no subtables. Index 1 -> 'arab', a Type 6 lookup whose
        /// ChainingContextualSubstFormat3 matches GlyphA and substitutes it via a nested Type 1
        /// lookup. Index 1 is last, so it wins a tag-only map regardless of active script.
        /// </summary>
        private static GsubTable BuildChainingContextualGsubTable()
        {
            var singleSubst = new SingleSubstSubTableFormat2
            {
                Coverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { GlyphA } },
                SubstituteGlyphIDs = new ushort[] { SubstituteGlyph }
            };

            var contextualRule = new ChainingContextualSubstFormat3
            {
                InputCoverages = new List<CoverageTable>
                {
                    new CoverageTableFormat1 { GlyphArray = new ushort[] { GlyphA } }
                },
                SubstLookupRecords = new List<SubstLookupRecord>
                {
                    new SubstLookupRecord { SequenceIndex = 0, LookupListIndex = 2 } // targets the Type 1 lookup at index 2
                }
            };

            var latinLookup = new LookupTable { LookupType = 6, SubTables = new List<FontTableElement>() };
            var arabLookup = new LookupTable
            {
                LookupType = 6,
                SubTables = new List<FontTableElement> { contextualRule }
            };
            var singleSubstLookup = new LookupTable
            {
                LookupType = 1,
                SubTables = new List<FontTableElement> { singleSubst }
            };

            return new GsubTable
            {
                ScriptList = BuildTwoScriptList(),
                FeatureList = BuildTwoFeatureList("liga"),
                LookupList = new LookupListTable
                {
                    // Index 0 = latinLookup (feature index 0 -> LookupListIndices [0]).
                    // Index 1 = arabLookup (feature index 1 -> LookupListIndices [1], matches
                    // BuildTwoFeatureList's fixed {0}/{1} pattern).
                    // Index 2 = singleSubstLookup, referenced only by arabLookup's own
                    // SubstLookupRecord.LookupListIndex, not by any FeatureRecord.
                    Lookups = new List<LookupTable> { latinLookup, arabLookup, singleSubstLookup }
                }
            };
        }

        [TestMethod]
        public void SingleSubstitution_ExtensionWrapped_IsApplied()
        {
            var font = LoadHostFont();
            font.AddOrReplaceTable(BuildExtensionWrappedSingleSubstGsubTable());

            var processor = new SingleSubstitutionProcessor(font);

            var glyphs = new List<ShapedGlyph>
            {
                new ShapedGlyph { GlyphId = GlyphA, BaseAdvance = 500, XAdvance = 500 }
            };

            var result = processor.ApplySubstitutions(
                glyphs, new List<string> { "smcp" }, "latn", null);

            Assert.AreEqual(
                SubstituteGlyph,
                result[0].GlyphId,
                "GlyphA must be substituted even though the \"smcp\" lookup is Extension "
                + "Substitution (Type 7) wrapping a SingleSubstSubTable, not a direct Type 1 lookup");
        }

        /// <summary>
        /// One "smcp" FeatureRecord -> a Type 7 (Extension) lookup wrapping a Type 1
        /// SingleSubstSubTable substituting GlyphA -> SubstituteGlyph. No ScriptList: this test
        /// is about extension unwrapping, not script filtering.
        /// </summary>
        private static GsubTable BuildExtensionWrappedSingleSubstGsubTable()
        {
            var singleSubst = new SingleSubstSubTableFormat2
            {
                Coverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { GlyphA } },
                SubstituteGlyphIDs = new ushort[] { SubstituteGlyph }
            };

            var extensionWrapper = new ExtensionSubstSubTable
            {
                ExtensionLookupType = 1,
                ExtendedSubTable = singleSubst
            };

            var lookup = new LookupTable
            {
                LookupType = 7,
                SubTables = new List<FontTableElement> { extensionWrapper }
            };

            var featureList = new FeatureListTable
            {
                FeatureRecords = new List<FeatureRecord>
                {
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("smcp"),
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

        /// <summary>
        /// Two "smcp" FeatureRecords: index 0 -> 'latn', a lookup with no subtables. Index 1 ->
        /// 'arab', a SingleSubstSubTable substituting GlyphA -> SubstituteGlyph. Index 1 is last,
        /// so it wins a tag-only map regardless of active script.
        /// </summary>
        private static GsubTable BuildSingleSubstGsubTable()
        {
            var arabSubst = new SingleSubstSubTableFormat2
            {
                Coverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { GlyphA } },
                SubstituteGlyphIDs = new ushort[] { SubstituteGlyph }
            };

            var latinLookup = new LookupTable { LookupType = 1, SubTables = new List<FontTableElement>() };
            var arabLookup = new LookupTable
            {
                LookupType = 1,
                SubTables = new List<FontTableElement> { arabSubst }
            };

            return new GsubTable
            {
                ScriptList = BuildTwoScriptList(),
                FeatureList = BuildTwoFeatureList("smcp"),
                LookupList = new LookupListTable { Lookups = new List<LookupTable> { latinLookup, arabLookup } }
            };
        }

        /// <summary>
        /// Two FeatureRecords sharing the given tag: index 0 references lookup 0, index 1
        /// references lookup 1 (or, for the chaining-contextual table, lookup 2 - see
        /// BuildChainingContextualGsubTable's own lookup list).
        /// </summary>
        private static FeatureListTable BuildTwoFeatureList(string tag)
        {
            return new FeatureListTable
            {
                FeatureRecords = new List<FeatureRecord>
                {
                    new FeatureRecord
                    {
                        FeatureTag = new Tag(tag),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 0 } }
                    },
                    new FeatureRecord
                    {
                        FeatureTag = new Tag(tag),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 1 } }
                    }
                }
            };
        }

        /// <summary>
        /// 'latn' -> feature index 0 only. 'arab' -> feature index 1 only. Shared by all three
        /// synthetic tables in this file; each table's own LookupList is ordered so that feature
        /// index 1 always resolves to LookupListIndices [1], per BuildTwoFeatureList.
        /// </summary>
        private static ScriptListTable BuildTwoScriptList()
        {
            return new ScriptListTable
            {
                ScriptRecords = new List<ScriptRecord>
                {
                    new ScriptRecord
                    {
                        ScriptTag = new Tag("latn"),
                        ScriptTable = new ScriptTable
                        {
                            DefaultLangSys = new LangSysTable
                            {
                                RequiredFeatureIndex = 0xFFFF,
                                FeatureIndices = new ushort[] { 0 }
                            }
                        }
                    },
                    new ScriptRecord
                    {
                        ScriptTag = new Tag("arab"),
                        ScriptTable = new ScriptTable
                        {
                            DefaultLangSys = new LangSysTable
                            {
                                RequiredFeatureIndex = 0xFFFF,
                                FeatureIndices = new ushort[] { 1 }
                            }
                        }
                    }
                }
            };
        }
    }
}