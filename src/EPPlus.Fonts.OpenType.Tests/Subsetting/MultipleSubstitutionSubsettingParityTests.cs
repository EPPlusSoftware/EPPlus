/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           GSUB Multiple Substitution subsetting parity test
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

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    /// <summary>
    /// Parity test analogous to the cmap format-14 subsetting parity test: verifies that shaping
    /// against the FULL font (the measurement path) and shaping against the ROUND-TRIPPED SUBSET
    /// font (the render path - what actually ends up embedded in the PDF) agree on a GSUB Multiple
    /// Substitution (Lookup Type 2).
    ///
    /// "Round-tripped" means serialized to bytes and reloaded - exactly like a font that has gone
    /// through PDF embedding. This exercises MultipleSubstHandler's discovery/rewrite AND
    /// GsubTableLoader's parsing of Lookup Type 2, both of which are new.
    /// </summary>
    [TestClass]
    public class MultipleSubstitutionSubsettingParityTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private OpenTypeFont LoadFullFontWithSyntheticMultipleSubst(out ushort glyphIdX, out ushort glyphIdA, out ushort glyphIdB)
        {
            var font = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Regular, true);

            Assert.IsTrue(font.CmapTable.TryGetGlyphId('X', out glyphIdX));
            Assert.IsTrue(font.CmapTable.TryGetGlyphId('A', out glyphIdA));
            Assert.IsTrue(font.CmapTable.TryGetGlyphId('B', out glyphIdB));


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
        public void Shape_RoundTrippedSubsetFont_StillExpandsMultipleSubstitution()
        {
            const string text = "X";

            var fullFont = LoadFullFontWithSyntheticMultipleSubst(out ushort glyphIdX, out ushort glyphIdA, out ushort glyphIdB);

            // --- Measurement path: shape against the full font ---
            var fullShaper = new TextShaper(TestFolderEngine, fullFont);
            var fullShaped = fullShaper.Shape(text);

            Assert.AreEqual(2, fullShaped.Glyphs.Length, "Sanity check: the full font must expand 'X' into two glyphs.");
            Assert.AreEqual(glyphIdA, fullShaped.Glyphs[0].GlyphId);
            Assert.AreEqual(glyphIdB, fullShaped.Glyphs[1].GlyphId);

            // --- Render path: subset for exactly this text, then round-trip (serialize + reload) -
            // exactly what happens when the subset is embedded in, and later read back from, a PDF. ---
            var subsetFont = fullFont.CreateSubset(text);
            var subsetBytes = subsetFont.Serialize();
            var roundTrippedSubsetFont = new OpenTypeFont(subsetBytes);

            var subsetShaper = new TextShaper(TestFolderEngine, roundTrippedSubsetFont);
            var subsetShaped = subsetShaper.Shape(text);

            Assert.AreEqual(2, subsetShaped.Glyphs.Length,
                "The round-tripped SUBSET font must still expand 'X' into two glyphs - today the Multiple " +
                "Substitution lookup isn't discovered/preserved by subsetting, and Lookup Type 2 isn't even " +
                "parsed back when the subset font is reloaded, so this silently collapses to one glyph.");

            // Glyph ids are renumbered by subsetting, so compare relative identity, not raw ids:
            // the two resulting glyphs in the subset must still be DIFFERENT from each other, and
            // different from whatever "X" alone would resolve to without the substitution.
            Assert.AreNotEqual(subsetShaped.Glyphs[0].GlyphId, subsetShaped.Glyphs[1].GlyphId,
                "The two expanded glyphs must remain distinct after subsetting.");
        }
    }
}