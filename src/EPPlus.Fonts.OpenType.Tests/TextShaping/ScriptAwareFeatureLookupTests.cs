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
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1;
using EPPlus.Fonts.OpenType.TextShaping.Positioning;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    /// <summary>
    /// SingleAdjustmentProvider.BuildFeatureMap maps a feature TAG (e.g. "kern") straight to
    /// a list of subtables, with no notion of script. When several FeatureRecords share the same
    /// tag but are reachable through different scripts in ScriptList, the map keeps only the last
    /// one it iterates - so a lookup that should only ever fire for one script (here 'arab') can
    /// silently answer for another (here 'latn').
    ///
    /// This is the exact shape of a real defect found in Calibri Bold, where a SinglePos
    /// tagged "kern" and reachable only via 'arab' gave 'period' a YPlacement of 168 in Latin
    /// text. Reproducing it against the real Calibri Bold font is not reliable: subsetting
    /// collapses FeatureList down to only what was referenced, so the full, unsubsetted font has
    /// a different FeatureList shape and the same tag ends up resolving to different data. This
    /// test instead builds the minimal GposTable structure directly, so the defect is demonstrated
    /// against the mapping logic itself rather than against a specific font file.
    /// </summary>
    [TestClass]
    public class ScriptAwareFeatureLookupTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const ushort PeriodGlyphId = 5;

        [TestMethod]
        public void SingleAdjustment_ForLatinScript_DoesNotPickUpArabicOnlyLookup()
        {
            // BIZUDGothic has no GPOS table at all, so OpenTypeFont's internal GposTableLoader is
            // null and the GposTable getter falls through to the local table cache below. Any
            // font that already HAS a GPOS table would ignore the injected table: the getter
            // checks its loader first and only falls back to the cache when there is none.
            //
            // ignoreCache: true is required. The normal cached path marks the font read-only
            // (FontStore.LoadFont) so it can be shared safely across tests; AddOrReplaceTable
            // throws on a read-only instance.
            var font = TestFolderEngine.LoadFont("BIZUDGothic", FontSubFamily.Regular, ignoreCache: true);
            Assert.IsNotNull(font, "BIZUDGothic must be present in the test font folder");

            font.AddOrReplaceTable(BuildSyntheticGposTable());

            var provider = new SingleAdjustmentProvider(font);

            // "latn" is what ShapingOptions.Default/.Fast/.Full all pass. Previously,
            // TryGetAdjustment had no script parameter at all and always searched every
            // FeatureRecord tagged "kern"; now it must restrict to the script given here.
            bool found = provider.TryGetAdjustment(
                PeriodGlyphId,
                new List<string> { "kern" },
                script: "latn",
                language: null,
                out var value);

            if (found)
            {
                Assert.AreEqual(
                    0,
                    (int)(value?.YPlacement ?? 0),
                    "'period' must not receive YPlacement from the lookup that is only reachable "
                    + "through the 'arab' script's LangSys - SingleAdjustmentProvider has no "
                    + "script filter yet, so the tag-only map picks it up regardless of script.");
            }
        }

        /// <summary>
        /// Builds a GposTable with two "kern" FeatureRecords sharing the tag but reachable
        /// through different scripts:
        ///   - lookup 0 (SinglePos, YPlacement 168 on PeriodGlyphId) reachable only via 'arab'
        ///   - lookup 1 (SinglePos, no adjustment relevant to PeriodGlyphId) reachable via 'latn'
        /// FeatureRecords are ordered so the 'arab' record comes LAST, which is what makes it win
        /// the tag-keyed dictionary in BuildFeatureMap regardless of which script is active.
        /// </summary>
        private static GposTable BuildSyntheticGposTable()
        {
            var arabSubtable = new SinglePosSubTableFormat1
            {
                SubtableFormat = 1,
                ValueFormat = 0x0002, // YPlacement only
                Coverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { PeriodGlyphId } },
                Value = new ValueRecord { YPlacement = 168 }
            };

            var latinSubtable = new SinglePosSubTableFormat1
            {
                SubtableFormat = 1,
                ValueFormat = 0x0004, // XAdvance only - a plausible, unrelated latn kern adjustment
                Coverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { 42 } }, // not 'period'
                Value = new ValueRecord { XAdvance = -30 }
            };

            var lookupList = new LookupListTable
            {
                Lookups = new List<LookupTable>
                {
                    new LookupTable { LookupType = 1, SubTables = new List<FontTableElement> { arabSubtable } },
                    new LookupTable { LookupType = 1, SubTables = new List<FontTableElement> { latinSubtable } }
                }
            };

            var featureList = new FeatureListTable
            {
                FeatureRecords = new List<FeatureRecord>
                {
                    // Index 0: 'latn' record -> lookup 1 (does not cover 'period').
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("kern"),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 1 } }
                    },
                    // Index 1: 'arab' record -> lookup 0 (covers 'period', YPlacement 168).
                    // Placed last so it is the one that wins the OLD, tag-only dictionary,
                    // regardless of which script is actually active - that overwrite is the bug.
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("kern"),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 0 } }
                    }
                }
            };

            var scriptList = new ScriptListTable
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
                                FeatureIndices = new ushort[] { 0 } // only the latn record
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
                                FeatureIndices = new ushort[] { 1 } // only the arab record
                            }
                        }
                    }
                }
            };

            return new GposTable
            {
                ScriptList = scriptList,
                FeatureList = featureList,
                LookupList = lookupList
            };
        }
    }
}