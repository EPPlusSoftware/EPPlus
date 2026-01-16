/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/22/2025         EPPlus Software AB           Subsetting edge case tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using EPPlus.Fonts.OpenType.TextShaping;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class SubsettingEdgeCasesTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Subset_EmptyString_ShouldThrow()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Subset_NullArray_ShouldThrow()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset((char[])null);
        }

        [TestMethod]
        public void Subset_SingleChar_ShouldHaveMinimalGlyphs()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("a");

            SaveFont("edge_single_char.ttf", subset);

            Assert.IsNotNull(subset);
            Assert.IsTrue(subset.MaxpTable.numGlyphs >= 3, 
                "Should have at least .notdef, space, a");
        }

        [TestMethod]
        public void Subset_LargeText_ShouldCompleteQuickly()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var allLatinChars = Enumerable.Range(32, 95).Select(i => (char)i).ToArray();

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var subset = font.CreateSubset(allLatinChars);
            sw.Stop();

            SaveFont("edge_all_latin.ttf", subset);

            Assert.IsTrue(sw.ElapsedMilliseconds < 5000,
                $"Subsetting took too long: {sw.ElapsedMilliseconds}ms");
        }

        [TestMethod]
        public void Subset_DuplicateChars_ShouldDedup()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var subset1 = font.CreateSubset("aaa");
            var subset2 = font.CreateSubset("a");

            Assert.AreEqual(subset1.MaxpTable.numGlyphs, subset2.MaxpTable.numGlyphs,
                "Duplicate characters should result in same glyph count");
        }

        [TestMethod]
        public void Subset_AllGlyphs_ShouldBeSimilarSizeToOriginal()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            // Get ALL characters from cmap
            var allChars = new HashSet<char>();
            foreach (var subtable in font.CmapTable.SubTables)
            {
                var mappings = subtable.GetGlyphMappings().CharCodeToGlyphIndex;
                foreach (var codepoint in mappings.Keys)
                {
                    if (codepoint <= 0xFFFF)
                        allChars.Add((char)codepoint);
                }
            }

            var subset = font.CreateSubset(allChars);

            SaveFont("edge_full_subset.ttf", subset);

            // Subset with ALL glyphs should be similar size to original
            var originalSize = font.Serialize().Length;
            var subsetSize = subset.Serialize().Length;

            double ratio = (double)subsetSize / originalSize;
            Assert.IsTrue(ratio > 0.8,
                $"Full subset unexpectedly small: {ratio:P0} of original");
        }

        [TestMethod]
        public void Subset_PreservesRobotoKerning()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var builder = new SubsetFontBuilder();

            // Act
            var subset = font.CreateSubset("AV");
            var shaper = new TextShaper(subset);
            var result = shaper.Shape("AV");

            // Assert
            Assert.IsTrue(result.Glyphs[0].XAdvance < font.HmtxTable.GetAdvanceWidth(
                font.CmapTable.MapCharToGlyph('A')),
                "Subset should preserve A-V kerning");
        }
    }
}