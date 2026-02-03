/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPPlus.Fonts.OpenType.TextShaping;
using System.Linq;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    [TestClass]
    public class ChainingContextualSubstitutionTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region Roboto ffi Ligature Tests (Type 6 Contextual)

        [TestMethod]
        public void ChainingContextual_Roboto_FfiLigature_Office()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.Shape("office");

            // Assert - Expected: 'o' + 'ffi' ligature + 'c' + 'e' = 4 glyphs
            Assert.AreEqual(4, result.Glyphs.Length, "Should have 4 glyphs: o, ffi, c, e");
            Assert.AreEqual(3, result.Glyphs[1].CharCount, "ffi ligature should represent 3 characters");
        }

        [TestMethod]
        public void ChainingContextual_Roboto_FfiLigature_AtStart()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act - ffi at the beginning of text (no backtrack context)
            var result = shaper.Shape("fficer");

            // Assert - Expected: 'ffi' ligature + 'c' + 'e' + 'r' = 4 glyphs
            Assert.AreEqual(4, result.Glyphs.Length, "Should have 4 glyphs: ffi, c, e, r");
            Assert.AreEqual(3, result.Glyphs[0].CharCount, "ffi ligature should represent 3 characters");
        }

        [TestMethod]
        public void ChainingContextual_Roboto_FfiLigature_AtEnd()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act - ffi at the end of text (no lookahead context)
            var result = shaper.Shape("offi");

            // Assert - Expected: 'o' + 'ffi' ligature = 2 glyphs
            Assert.AreEqual(2, result.Glyphs.Length, "Should have 2 glyphs: o, ffi");
            Assert.AreEqual(3, result.Glyphs[1].CharCount, "ffi ligature should represent 3 characters");
        }

        [TestMethod]
        public void ChainingContextual_Roboto_MultipleFfiLigatures()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act - Multiple ffi sequences in same text
            var result = shaper.Shape("office officer");

            // Assert - Expected: 'o' + 'ffi' + 'c' + 'e' + ' ' + 'o' + 'ffi' + 'c' + 'e' + 'r' = 10 glyphs
            Assert.AreEqual(10, result.Glyphs.Length);
            Assert.AreEqual(3, result.Glyphs[1].CharCount, "First ffi ligature");
            Assert.AreEqual(3, result.Glyphs[6].CharCount, "Second ffi ligature");
        }

        #endregion

        #region Type 6 vs Type 4 Interaction

        [TestMethod]
        public void ChainingContextual_Roboto_Type6BeforeType4_CorrectOrder()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("office fit");
            var shaper = new TextShaper(subset);

            // Act - Text with both ffi (Type 6) and fi (Type 4) ligatures
            var result = shaper.Shape("office fit");

            // Assert - Expected: 'o' + 'ffi' + 'c' + 'e' + ' ' + 'fi' + 't' = 7 glyphs
            Assert.AreEqual(7, result.Glyphs.Length);
            Assert.AreEqual(3, result.Glyphs[1].CharCount, "ffi from Type 6 contextual");
            Assert.AreEqual(2, result.Glyphs[5].CharCount, "fi from Type 4 simple");
        }

        #endregion

        #region Metrics Validation

        [TestMethod]
        public void ChainingContextual_Roboto_FfiLigature_HasCorrectMetrics()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.Shape("ffi");

            // Assert
            Assert.AreEqual(1, result.Glyphs.Length, "Should be single ffi ligature");

            var ffiGlyph = result.Glyphs[0];
            Assert.IsTrue(ffiGlyph.XAdvance > 0, "ffi ligature should have positive advance width");
            Assert.AreEqual(0, ffiGlyph.YAdvance, "Horizontal text should have zero Y advance");
            Assert.AreEqual(0, ffiGlyph.ClusterIndex, "Should start at cluster 0");
            Assert.AreEqual(3, ffiGlyph.CharCount, "Should represent 3 characters");
        }

        #endregion
    }
}