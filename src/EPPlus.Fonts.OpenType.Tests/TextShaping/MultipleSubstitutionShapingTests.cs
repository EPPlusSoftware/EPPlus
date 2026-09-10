/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           GSUB Multiple Substitution shaping tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    /// <summary>
    /// Verifies that TextShaper actually applies GSUB Lookup Type 2 (Multiple Substitution) -
    /// today the engine has handlers for Type 1 (single), Type 4 (ligature) and Type 6 (chaining
    /// contextual), but nothing expands one input glyph into several output glyphs, which is
    /// what most real-world "ccmp" decomposition relies on.
    ///
    /// The substitution data is injected synthetically into a real, otherwise-unmodified test
    /// font, so the tests don't depend on any specific font shipping a Type 2 lookup for a
    /// character we control. It's tagged "liga" and appended to the font's own existing "liga"
    /// feature so it activates through ShapingOptions.Default without touching any public
    /// GsubFeature enum - this test is only about the lookup TYPE working, not about exposing a
    /// new feature tag publicly.
    /// </summary>
    [TestClass]
    public class MultipleSubstitutionShapingTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        /// <summary>
        /// Loads a private (non-cached) Roboto instance and appends a synthetic Multiple
        /// Substitution lookup - triggered by 'X' - to the font's existing "liga" feature. 'X'
        /// decomposes into the glyphs for 'A' and 'B' (two real, distinct, existing glyphs -
        /// not made up ids), so the effect is trivially visible: shaping "X" should produce the
        /// same two glyphs as shaping "AB", not the single glyph for 'X'.
        /// </summary>
        private OpenTypeFont LoadFontWithSyntheticMultipleSubst(out ushort glyphIdX, out ushort glyphIdA, out ushort glyphIdB)
        {
            var font = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Regular, true);

            Assert.IsTrue(font.CmapTable.TryGetGlyphId('X', out glyphIdX), "Test font is expected to contain 'X'.");
            Assert.IsTrue(font.CmapTable.TryGetGlyphId('A', out glyphIdA), "Test font is expected to contain 'A'.");
            Assert.IsTrue(font.CmapTable.TryGetGlyphId('B', out glyphIdB), "Test font is expected to contain 'B'.");


            var multiSubst = new MultipleSubstSubTable
            {
                SubtableFormat = 1,
                Coverage = CoverageTableFormat2.CreateCoverageFormat2(new List<ushort> { glyphIdX }),
                Sequences = new List<ushort[]> { new[] { glyphIdA, glyphIdB } }
            };

            var newLookup = new LookupTable
            {
                LookupType = 2,
                LookupFlag = 0,
                SubTables = new List<FontTableElement> { multiSubst }
            };

            int newLookupIndex = font.GsubTable.LookupList.Lookups.Count;
            font.GsubTable.LookupList.Lookups.Add(newLookup);

            AppendLookupToActiveLigaFeature(font, newLookupIndex);

            return font;
        }

        /// <summary>
        /// Returns the index of a "liga" FeatureRecord that is actually REACHABLE from the
        /// "latn" script, and appends <paramref name="lookupIndex"/> to it.
        ///
        /// This matters: Roboto defines "liga" three times (once per script grouping), and only
        /// ONE of those FeatureRecords is listed in the latn LangSys. Picking the first "liga"
        /// by tag alone lands on a record the script-aware feature resolver correctly filters
        /// out, so the injected lookup would never run - the test would fail for a reason that
        /// has nothing to do with the code under test.
        /// </summary>
        private static void AppendLookupToActiveLigaFeature(OpenTypeFont font, int lookupIndex)
        {
            var activeIndices = ScriptFeatureResolver.GetActiveFeatureIndices(
                font.GsubTable.ScriptList, "latn", null);
            Assert.IsNotNull(activeIndices, "Test font is expected to have a ScriptList.");

            var records = font.GsubTable.FeatureList.FeatureRecords;
            int ligaIndex = -1;
            for (int i = 0; i < records.Count; i++)
            {
                if (records[i].FeatureTag.Value == "liga" && activeIndices.Contains(i))
                {
                    ligaIndex = i;
                    break;
                }
            }

            Assert.IsTrue(ligaIndex >= 0, "Test font is expected to define a 'liga' feature reachable from the latn script.");

            var featureTable = records[ligaIndex].FeatureTable;
            var oldIndices = featureTable.LookupListIndices ?? new ushort[0];
            var newIndices = new ushort[oldIndices.Length + 1];
            oldIndices.CopyTo(newIndices, 0);
            newIndices[oldIndices.Length] = (ushort)lookupIndex;
            featureTable.LookupListIndices = newIndices;
            featureTable.LookupCount = (ushort)newIndices.Length;
        }

        [TestMethod]
        public void Shape_GlyphWithMultipleSubstitution_ExpandsIntoItsSequence()
        {
            var font = LoadFontWithSyntheticMultipleSubst(out ushort glyphIdX, out ushort glyphIdA, out ushort glyphIdB);
            var shaper = new TextShaper(TestFolderEngine, font);

            var shaped = shaper.Shape("X");

            Assert.AreEqual(2, shaped.Glyphs.Length, "'X' should expand into its two-glyph substitution sequence.");
            Assert.AreEqual(glyphIdA, shaped.Glyphs[0].GlyphId);
            Assert.AreEqual(glyphIdB, shaped.Glyphs[1].GlyphId);
        }

        [TestMethod]
        public void Shape_GlyphWithMultipleSubstitution_PreservesTotalCharCountAndClusterIndex()
        {
            var font = LoadFontWithSyntheticMultipleSubst(out ushort glyphIdX, out ushort glyphIdA, out ushort glyphIdB);
            var shaper = new TextShaper(TestFolderEngine, font);

            var shaped = shaper.Shape("X");

            int totalCharCount = shaped.Glyphs[0].CharCount + shaped.Glyphs[1].CharCount;
            Assert.AreEqual(1, totalCharCount, "The single source character must still be fully accounted for across the expanded glyphs.");
            Assert.AreEqual(0, shaped.Glyphs[0].ClusterIndex);
            Assert.AreEqual(0, shaped.Glyphs[1].ClusterIndex, "Both expanded glyphs map back to the same source character position.");
        }

        [TestMethod]
        public void Shape_GlyphWithMultipleSubstitution_EachGlyphGetsItsOwnAdvanceWidth()
        {
            var font = LoadFontWithSyntheticMultipleSubst(out ushort glyphIdX, out ushort glyphIdA, out ushort glyphIdB);
            var shaper = new TextShaper(TestFolderEngine, font);

            var shaped = shaper.Shape("X");
            var plainAB = shaper.Shape("AB");

            Assert.AreEqual(plainAB.Glyphs[0].BaseAdvance, shaped.Glyphs[0].BaseAdvance,
                "Each expanded glyph should carry its own real advance width, not a shared/zero one.");
            Assert.AreEqual(plainAB.Glyphs[1].BaseAdvance, shaped.Glyphs[1].BaseAdvance);
        }

        [TestMethod]
        public void Shape_GlyphWithoutTrigger_IsUnaffected()
        {
            var font = LoadFontWithSyntheticMultipleSubst(out ushort glyphIdX, out ushort glyphIdA, out ushort glyphIdB);
            var shaper = new TextShaper(TestFolderEngine, font);

            var shaped = shaper.Shape("A");

            Assert.AreEqual(1, shaped.Glyphs.Length, "A character with no Multiple Substitution entry must be completely unaffected.");
            Assert.AreEqual(glyphIdA, shaped.Glyphs[0].GlyphId);
            Assert.AreEqual(1, shaped.Glyphs[0].CharCount);
        }
    }
}