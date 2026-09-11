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
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using EPPlus.Fonts.OpenType.TextShaping.Kerning;
using EPPlus.Fonts.OpenType.TextShaping.Positioning;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    /// <summary>
    /// Same synthetic-GposTable technique as ScriptAwareFeatureLookupTests, applied to
    /// GposKerningProvider and MarkToBaseProvider: two FeatureRecords share a tag ("kern" /
    /// "mark") but are reachable through different scripts, and the 'arab'-only record is placed
    /// last so it wins any tag-only mapping regardless of which script is actually active.
    /// </summary>
    [TestClass]
    public class ScriptAwareKerningAndMarkTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const ushort GlyphA = 5;
        private const ushort GlyphB = 6;
        private const ushort MarkGlyph = 7;

        [TestMethod]
        public void Kerning_ForLatinScript_DoesNotPickUpArabicOnlyPair()
        {
            var font = TestFolderEngine.LoadFont("BIZUDGothic", FontSubFamily.Regular, ignoreCache: true);
            Assert.IsNotNull(font, "BIZUDGothic must be present in the test font folder");

            font.AddOrReplaceTable(BuildKerningGposTable());

            var provider = new GposKerningProvider(font.GposTable);

            short kerning = provider.GetKerning(GlyphA, GlyphB, "latn", null);

            Assert.AreEqual(
                0,
                (int)kerning,
                "A+B must not receive kerning from the pair that is only reachable through the "
                + "'arab' script's LangSys");
        }

        [TestMethod]
        public void MarkPositioning_ForLatinScript_DoesNotPickUpArabicOnlyOffset()
        {
            var font = TestFolderEngine.LoadFont("BIZUDGothic", FontSubFamily.Regular, ignoreCache: true);
            Assert.IsNotNull(font, "BIZUDGothic must be present in the test font folder");

            font.AddOrReplaceTable(BuildMarkGposTable());

            var provider = new MarkToBaseProvider(font);

            var glyphs = new List<ShapedGlyph>
            {
                new ShapedGlyph { GlyphId = GlyphA, XAdvance = 500, BaseAdvance = 500 },
                new ShapedGlyph { GlyphId = MarkGlyph, XAdvance = 0, BaseAdvance = 0 }
            };

            provider.ApplyMarkPositioning(glyphs, "latn", null);

            Assert.AreEqual(
                0,
                (int)glyphs[1].YOffset,
                "the mark must not receive an offset from the MarkToBase subtable that is only "
                + "reachable through the 'arab' script's LangSys");
        }

        /// <summary>
        /// Two "kern" FeatureRecords: index 0 -> 'latn', covers a pair NOT involving A/B.
        /// Index 1 -> 'arab', covers A+B with a large adjustment. Index 1 is last, so it is the
        /// one that would win a tag-only (script-unaware) map regardless of active script.
        /// </summary>
        private static GposTable BuildKerningGposTable()
        {
            var arabPair = new PairPosSubTableFormat1
            {
                Coverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { GlyphA } },
                PairSets = new List<PairSet>
                {
                    new PairSet
                    {
                        PairValueRecords = new List<PairValueRecord>
                        {
                            new PairValueRecord
                            {
                                SecondGlyph = GlyphB,
                                Value1 = new ValueRecord { XAdvance = -200 }
                            }
                        }
                    }
                }
            };

            var latinPair = new PairPosSubTableFormat1
            {
                Coverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { 42 } },
                PairSets = new List<PairSet>
                {
                    new PairSet
                    {
                        PairValueRecords = new List<PairValueRecord>
                        {
                            new PairValueRecord
                            {
                                SecondGlyph = 43,
                                Value1 = new ValueRecord { XAdvance = -10 }
                            }
                        }
                    }
                }
            };

            var lookupList = new LookupListTable
            {
                Lookups = new List<LookupTable>
                {
                    new LookupTable { LookupType = 2, SubTables = new List<FontTableElement> { latinPair } },
                    new LookupTable { LookupType = 2, SubTables = new List<FontTableElement> { arabPair } }
                }
            };

            var featureList = new FeatureListTable
            {
                FeatureRecords = new List<FeatureRecord>
                {
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("kern"),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 0 } }
                    },
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("kern"),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 1 } }
                    }
                }
            };

            return new GposTable
            {
                ScriptList = BuildTwoScriptList(),
                FeatureList = featureList,
                LookupList = lookupList
            };
        }

        /// <summary>
        /// Two "mark" FeatureRecords: index 0 -> 'latn', covers a base/mark pair that never
        /// appears in the test glyph sequence. Index 1 -> 'arab', covers GlyphA/MarkGlyph with a
        /// large YOffset. Index 1 is last, so it is the one that would win a tag-only map.
        /// </summary>
        private static GposTable BuildMarkGposTable()
        {
            var arabMark = new MarkToBaseSubTableFormat1
            {
                MarkClassCount = 1,
                MarkCoverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { MarkGlyph } },
                BaseCoverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { GlyphA } },
                MarkArray = new MarkArray
                {
                    MarkCount = 1,
                    Records = new[]
                    {
                        new MarkRecord { MarkClass = 0, MarkAnchor = new AnchorTable { XCoordinate = 0, YCoordinate = 0 } }
                    }
                },
                BaseArray = new BaseArray
                {
                    BaseCount = 1,
                    Records = new[]
                    {
                        new BaseRecord
                        {
                            BaseAnchors = new[] { new AnchorTable { XCoordinate = 0, YCoordinate = 500 } }
                        }
                    }
                }
            };

            var latinMark = new MarkToBaseSubTableFormat1
            {
                MarkClassCount = 1,
                MarkCoverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { 44 } },
                BaseCoverage = new CoverageTableFormat1 { GlyphArray = new ushort[] { 45 } },
                MarkArray = new MarkArray
                {
                    MarkCount = 1,
                    Records = new[]
                    {
                        new MarkRecord { MarkClass = 0, MarkAnchor = new AnchorTable { XCoordinate = 0, YCoordinate = 0 } }
                    }
                },
                BaseArray = new BaseArray
                {
                    BaseCount = 1,
                    Records = new[]
                    {
                        new BaseRecord
                        {
                            BaseAnchors = new[] { new AnchorTable { XCoordinate = 0, YCoordinate = 10 } }
                        }
                    }
                }
            };

            var lookupList = new LookupListTable
            {
                Lookups = new List<LookupTable>
                {
                    new LookupTable { LookupType = 4, SubTables = new List<FontTableElement> { latinMark } },
                    new LookupTable { LookupType = 4, SubTables = new List<FontTableElement> { arabMark } }
                }
            };

            var featureList = new FeatureListTable
            {
                FeatureRecords = new List<FeatureRecord>
                {
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("mark"),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 0 } }
                    },
                    new FeatureRecord
                    {
                        FeatureTag = new Tag("mark"),
                        FeatureTable = new FeatureTable { LookupListIndices = new ushort[] { 1 } }
                    }
                }
            };

            return new GposTable
            {
                ScriptList = BuildTwoScriptList(),
                FeatureList = featureList,
                LookupList = lookupList
            };
        }

        /// <summary>
        /// 'latn' -> feature index 0 only. 'arab' -> feature index 1 only.
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