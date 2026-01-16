/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/16/2025         EPPlus Software AB           TextShaper tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    [TestClass]
    public class TextShaperTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region Basic Shaping Tests

        [TestMethod]
        public void Shape_EmptyString_ReturnsEmptyResult()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("");

            // Assert
            Assert.IsNotNull(shaped);
            Assert.AreEqual("", shaped.OriginalText);
            Assert.AreEqual(0, shaped.Glyphs.Length);
            Assert.AreEqual(0, shaped.TotalAdvanceWidth);
        }

        [TestMethod]
        public void Shape_NullString_ReturnsEmptyResult()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape(null);

            // Assert
            Assert.IsNotNull(shaped);
            Assert.AreEqual("", shaped.OriginalText);
            Assert.AreEqual(0, shaped.Glyphs.Length);
        }

        [TestMethod]
        public void Shape_SingleCharacter_ReturnsOneGlyph()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("A");

            // Assert
            Assert.IsNotNull(shaped);
            Assert.AreEqual("A", shaped.OriginalText);
            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.IsTrue(shaped.Glyphs[0].GlyphId > 0, "Should have valid glyph ID");
            Assert.IsTrue(shaped.Glyphs[0].XAdvance > 0, "Should have positive advance width");
        }

        [TestMethod]
        public void Shape_SimpleWord_ReturnsCorrectGlyphCount()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("Hello");

            // Assert
            Assert.IsNotNull(shaped);
            Assert.AreEqual("Hello", shaped.OriginalText);
            Assert.AreEqual(5, shaped.Glyphs.Length);
            Assert.IsTrue(shaped.TotalAdvanceWidth > 0);
        }

        [TestMethod]
        public void Shape_WithSpace_IncludesSpaceGlyph()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("A B");

            // Assert
            Assert.AreEqual(3, shaped.Glyphs.Length);
            Assert.IsTrue(shaped.Glyphs[1].XAdvance > 0, "Space should have advance width");
        }

        #endregion

        #region Kerning Tests

        [TestMethod]
        public void Shape_WithKerning_ReducesWidth()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var withKerning = shaper.Shape("WAVE", ShapingOptions.Default);
            var withoutKerning = shaper.Shape("WAVE", ShapingOptions.None);

            // Assert
            Assert.IsTrue(withKerning.TotalAdvanceWidth < withoutKerning.TotalAdvanceWidth,
                "Kerning should reduce width for 'WAVE'");
        }

        [TestMethod]
        public void Debug_GposKerningFormat()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            Assert.IsNotNull(font.GposTable, "Should have GPOS");

            // Find kern feature
            foreach (var featureRecord in font.GposTable.FeatureList.FeatureRecords)
            {
                if (featureRecord.FeatureTag.Value == "kern")
                {
                    Debug.WriteLine("Found 'kern' feature");

                    var feature = featureRecord.FeatureTable;
                    foreach (var lookupIndex in feature.LookupListIndices)
                    {
                        var lookup = font.GposTable.LookupList.Lookups[lookupIndex];
                        Debug.WriteLine($"Lookup type: {lookup.LookupType}");

                        foreach (var subtable in lookup.SubTables)
                        {
                            Debug.WriteLine($"Subtable type: {subtable.GetType().Name}");

                            if (subtable is PairPosSubTableFormat1 format1)
                            {
                                Debug.WriteLine($"  Format 1: {format1.PairSets.Count} pair sets");
                            }
                            else if (subtable is PairPosSubTableFormat2 format2)
                            {
                                Debug.WriteLine($"  Format 2: Class-based kerning");
                                Debug.WriteLine($"    ClassDef1 glyphs: {format2.ClassDef1?.GetType().Name}");
                                Debug.WriteLine($"    ClassDef2 glyphs: {format2.ClassDef2?.GetType().Name}");
                                Debug.WriteLine($"    Class1 count: {format2.Class1Count}");
                                Debug.WriteLine($"    Class2 count: {format2.Class2Count}");
                            }
                        }
                    }
                }
            }
        }

        [TestMethod]
        public void Shape_AVPair_HasNegativeKerning()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var withKerning = shaper.Shape("AV");
            var withoutKerning = shaper.Shape("AV", ShapingOptions.None);

            // Assert
            Assert.IsTrue(withKerning.TotalAdvanceWidth < withoutKerning.TotalAdvanceWidth,
                "A-V pair should have negative kerning");

            // Check that first glyph (A) has reduced advance
            Assert.IsTrue(withKerning.Glyphs[0].XAdvance < withoutKerning.Glyphs[0].XAdvance,
                "First glyph should have kerning applied");
        }

        [TestMethod]
        public void Shape_FastOption_StillAppliesKerning()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var fast = shaper.Shape("WAVE", ShapingOptions.Fast);
            var none = shaper.Shape("WAVE", ShapingOptions.None);

            // Assert
            Assert.IsTrue(fast.TotalAdvanceWidth < none.TotalAdvanceWidth,
                "Fast option should still apply kerning");
        }

        #endregion

        #region Measurement Tests

        [TestMethod]
        public void MeasureText_ReturnsPositiveWidth()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            int width = shaper.MeasureText("Hello");

            // Assert
            Assert.IsTrue(width > 0, "Width should be positive");
        }

        [TestMethod]
        public void MeasureTextInPoints_ReturnsReasonableValue()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            float width = shaper.MeasureTextInPoints("Hello", 12);

            // Assert
            Assert.IsTrue(width > 10, "Should be at least 10 points wide");
            Assert.IsTrue(width < 100, "Should be less than 100 points wide");
        }

        [TestMethod]
        public void MeasureTextInPixels_ScalesWithDpi()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            float width72 = shaper.MeasureTextInPixels("Hello", 12, 72);
            float width96 = shaper.MeasureTextInPixels("Hello", 12, 96);

            // Assert
            Assert.IsTrue(width96 > width72, "96 DPI should be wider than 72 DPI");
            Assert.AreEqual(96.0f / 72.0f, width96 / width72, 0.01, "Should scale proportionally");
        }

        [TestMethod]
        public void MeasureText_LargerFontSize_LargerWidth()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            float width12 = shaper.MeasureTextInPoints("Hello", 12);
            float width24 = shaper.MeasureTextInPoints("Hello", 24);

            // Assert
            Assert.IsTrue(width24 > width12 * 1.9, "24pt should be ~2x wider than 12pt");
            Assert.IsTrue(width24 < width12 * 2.1, "24pt should be ~2x wider than 12pt");
        }

        #endregion

        #region Glyph Properties Tests

        [TestMethod]
        public void Shape_GlyphsHaveClusterIndices()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("ABC");

            // Assert
            Assert.AreEqual(0, shaped.Glyphs[0].ClusterIndex);
            Assert.AreEqual(1, shaped.Glyphs[1].ClusterIndex);
            Assert.AreEqual(2, shaped.Glyphs[2].ClusterIndex);
        }

        [TestMethod]
        public void Shape_GlyphsHaveCharCount()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("ABC");

            // Assert
            foreach (var glyph in shaped.Glyphs)
            {
                Assert.AreEqual(1, glyph.CharCount, "Simple glyphs should have CharCount=1");
            }
        }

        [TestMethod]
        public void Shape_GlyphsHaveValidIds()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("Hello");

            // Assert
            foreach (var glyph in shaped.Glyphs)
            {
                Assert.IsTrue(glyph.GlyphId < font.MaxpTable.numGlyphs,
                    "Glyph ID should be within font bounds");
            }
        }

        #endregion

        #region Multi-line Tests

        [TestMethod]
        public void ShapeLines_SingleLine_ReturnsOneElement()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var lines = shaper.ShapeLines("Hello");

            // Assert
            Assert.AreEqual(1, lines.Length);
            Assert.AreEqual("Hello", lines[0].OriginalText);
        }

        [TestMethod]
        public void ShapeLines_TwoLinesWithLF_ReturnsTwoElements()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var lines = shaper.ShapeLines("Hello\nWorld");

            // Assert
            Assert.AreEqual(2, lines.Length);
            Assert.AreEqual("Hello", lines[0].OriginalText);
            Assert.AreEqual("World", lines[1].OriginalText);
        }

        [TestMethod]
        public void ShapeLines_TwoLinesWithCRLF_ReturnsTwoElements()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var lines = shaper.ShapeLines("Hello\r\nWorld");

            // Assert
            Assert.AreEqual(2, lines.Length);
            Assert.AreEqual("Hello", lines[0].OriginalText);
            Assert.AreEqual("World", lines[1].OriginalText);
        }

        [TestMethod]
        public void ShapeLines_EmptyLine_PreservesEmptyLine()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var lines = shaper.ShapeLines("Hello\n\nWorld");

            // Assert
            Assert.AreEqual(3, lines.Length);
            Assert.AreEqual("Hello", lines[0].OriginalText);
            Assert.AreEqual("", lines[1].OriginalText);
            Assert.AreEqual("World", lines[2].OriginalText);
        }

        [TestMethod]
        public void MeasureLines_SingleLine_MatchesSingleMeasurement()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var metrics = shaper.MeasureLines("Hello", 12);
            float singleWidth = shaper.MeasureTextInPoints("Hello", 12);

            // Assert
            Assert.AreEqual(1, metrics.LineCount);
            Assert.AreEqual(singleWidth, metrics.Width, 0.01f);
        }

        [TestMethod]
        public void MeasureLines_TwoLines_WidthIsMaxOfBoth()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var metrics = shaper.MeasureLines("Hi\nHello", 12);
            float hiWidth = shaper.MeasureTextInPoints("Hi", 12);
            float helloWidth = shaper.MeasureTextInPoints("Hello", 12);

            // Assert
            Assert.AreEqual(2, metrics.LineCount);
            Assert.AreEqual(Math.Max(hiWidth, helloWidth), metrics.Width, 0.01f);
        }

        [TestMethod]
        public void MeasureLines_TwoLines_HeightIsDoubleLineHeight()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var metrics = shaper.MeasureLines("Hello\nWorld", 12);
            float lineHeight = shaper.GetLineHeightInPoints(12);

            // Assert
            Assert.AreEqual(2, metrics.LineCount);
            Assert.AreEqual(2 * lineHeight, metrics.Height, 0.01f);
        }

        #endregion

        #region Height Calculation Tests

        [TestMethod]
        public void GetLineHeightInPoints_ReturnsPositiveValue()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            float lineHeight = shaper.GetLineHeightInPoints(12);

            // Assert
            Assert.IsTrue(lineHeight > 0, "Line height should be positive");
            Assert.IsTrue(lineHeight > 10, "Line height should be reasonable for 12pt");
            Assert.IsTrue(lineHeight < 30, "Line height should be reasonable for 12pt");
        }

        [TestMethod]
        public void GetFontHeightInPoints_ReturnsPositiveValue()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            float fontHeight = shaper.GetFontHeightInPoints(12);

            // Assert
            Assert.IsTrue(fontHeight > 0, "Font height should be positive");
            Assert.IsTrue(fontHeight > 10, "Font height should be reasonable for 12pt");
            Assert.IsTrue(fontHeight < 20, "Font height should be reasonable for 12pt");
        }

        [TestMethod]
        public void GetLineHeight_IsGreaterThanFontHeight()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            float lineHeight = shaper.GetLineHeightInPoints(12);
            float fontHeight = shaper.GetFontHeightInPoints(12);

            // Assert
            Assert.IsTrue(lineHeight >= fontHeight,
                "Line height (with line gap) should be >= font height");
        }

        [TestMethod]
        public void GetLineHeight_ScalesWithFontSize()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            float height12 = shaper.GetLineHeightInPoints(12);
            float height24 = shaper.GetLineHeightInPoints(24);

            // Assert
            Assert.AreEqual(2.0f, height24 / height12, 0.01f,
                "Line height should scale linearly with font size");
        }

        #endregion

        #region ShapedText Properties Tests

        [TestMethod]
        public void ShapedText_GetWidthInPoints_MatchesMeasureText()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("Hello");
            float width1 = shaped.GetWidthInPoints(12, font.HeadTable.UnitsPerEm);
            float width2 = shaper.MeasureTextInPoints("Hello", 12);

            // Assert
            Assert.AreEqual(width1, width2, 0.01f, "Both methods should return same width");
        }

        [TestMethod]
        public void ShapedText_GetWidthInPixels_MatchesMeasureText()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("Hello");
            float width1 = shaped.GetWidthInPixels(12, 96, font.HeadTable.UnitsPerEm);
            float width2 = shaper.MeasureTextInPixels("Hello", 12, 96);

            // Assert
            Assert.AreEqual(width1, width2, 0.01f, "Both methods should return same width");
        }

        #endregion

        #region Edge Cases

        [TestMethod]
        public void Shape_OnlySpaces_ReturnsGlyphs()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("   ");

            // Assert
            Assert.AreEqual(3, shaped.Glyphs.Length);
            Assert.IsTrue(shaped.TotalAdvanceWidth > 0, "Spaces should have width");
        }

        [TestMethod]
        public void Shape_SpecialCharacters_HandlesGracefully()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("@#$%");

            // Assert
            Assert.AreEqual(4, shaped.Glyphs.Length);
            Assert.IsTrue(shaped.TotalAdvanceWidth > 0);
        }

        [TestMethod]
        public void Shape_Numbers_ReturnsCorrectGlyphs()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var shaped = shaper.Shape("12345");

            // Assert
            Assert.AreEqual(5, shaped.Glyphs.Length);
            Assert.IsTrue(shaped.TotalAdvanceWidth > 0);
        }

        #endregion

        #region Shaping with ligatures
        [TestMethod]
        public void Shape_FiLigature_CombinesTwoGlyphs()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var withLigatures = shaper.Shape("fi", ShapingOptions.Default);
            var withoutLigatures = shaper.Shape("fi", ShapingOptions.None);

            // Assert
            Assert.IsTrue(withLigatures.Glyphs.Length < withoutLigatures.Glyphs.Length,
                "Ligatures should combine glyphs (fi → 1 glyph instead of 2)");
        }

        [TestMethod]
        public void Shape_Office_HasFfiLigature()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.Shape("office");

            // Assert - "ffi" should be one ligature glyph with CharCount=3
            bool hasFfiLigature = result.Glyphs.Any(g => g.CharCount == 3);
            Assert.IsTrue(hasFfiLigature, "Should find ffi ligature in 'office'");
        }

        [TestMethod]
        public void Shape_Ligature_PreservesClusterIndex()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.Shape("afi");

            // Assert
            // Glyph[0]: 'a' at cluster 0
            // Glyph[1]: 'fi' ligature at cluster 1, represents 2 chars
            Assert.AreEqual(0, result.Glyphs[0].ClusterIndex);
            Assert.AreEqual(1, result.Glyphs[1].ClusterIndex);
            Assert.AreEqual(2, result.Glyphs[1].CharCount);
        }
        #endregion
    }
}