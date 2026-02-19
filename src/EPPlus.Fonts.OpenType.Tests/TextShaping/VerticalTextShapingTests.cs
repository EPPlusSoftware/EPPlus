/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           Vertical text shaping tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tests;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Drawing.Text;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    [TestClass]
    public class VerticalTextShapingTests : FontTestBase
    {
        public override TestContext TestContext { get; set; }

        [ClassInitialize]
        public static void ClassInitialize(TestContext ctx)
        {
            FontDirectoriesTestHelper.ClassInitialize(ctx);
        }

        #region ShapeVertical tests

        [TestMethod]
        public void ShapeVertical_CjkText_ReturnsOneGlyphPerCharacter()
        {
            // Arrange - BIZ UDGothic has vmtx and is a CJK font
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeVertical("日本語");

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(3, result.Glyphs.Length, "Expected one glyph per character");
        }

        [TestMethod]
        public void ShapeVertical_CjkText_GlyphsHavePositiveYAdvance()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeVertical("日本語");

            // Assert - all glyphs must have a positive YAdvance (sourced from vmtx)
            foreach (var glyph in result.Glyphs)
            {
                Assert.IsTrue(glyph.YAdvance > 0,
                    $"Glyph {glyph.GlyphId} has YAdvance {glyph.YAdvance}, expected > 0");
            }
        }

        [TestMethod]
        public void ShapeVertical_CjkText_TotalAdvanceHeightIsPositive()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeVertical("テスト");

            // Assert
            Assert.IsTrue(result.TotalAdvanceHeight > 0,
                $"TotalAdvanceHeight should be positive, was {result.TotalAdvanceHeight}");
        }

        [TestMethod]
        public void ShapeVertical_EmptyString_ReturnsEmptyGlyphArray()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeVertical(string.Empty);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Glyphs.Length);
            Assert.AreEqual(string.Empty, result.OriginalText);
        }

        [TestMethod]
        public void ShapeVertical_ClusterIndexMatchesCharacterPosition()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);
            var text = "ABC";

            // Act
            var result = shaper.ShapeVertical(text);

            // Assert - ClusterIndex must map back to the correct character position
            Assert.AreEqual(3, result.Glyphs.Length);
            for (int i = 0; i < result.Glyphs.Length; i++)
            {
                Assert.AreEqual((ushort)i, result.Glyphs[i].ClusterIndex,
                    $"Glyph {i} has ClusterIndex {result.Glyphs[i].ClusterIndex}, expected {i}");
            }
        }

        #endregion

        #region ShapeLightVertical tests

        [TestMethod]
        public void ShapeLightVertical_CjkText_ReturnsOneEntryPerCharacter()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeLightVertical("日本語");

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(3, result.Length, "Expected one VerticalGlyphHeight per character");
        }

        [TestMethod]
        public void ShapeLightVertical_CjkText_YAdvanceMatchesShapeVertical()
        {
            // Arrange - ShapeLightVertical should produce identical YAdvance values to ShapeVertical
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);
            var text = "東京";

            // Act
            var full = shaper.ShapeVertical(text);
            var light = shaper.ShapeLightVertical(text);

            // Assert
            Assert.AreEqual(full.Glyphs.Length, light.Length);
            for (int i = 0; i < full.Glyphs.Length; i++)
            {
                Assert.AreEqual(full.Glyphs[i].YAdvance, light[i].YAdvance,
                    $"YAdvance mismatch at index {i}: ShapeVertical={full.Glyphs[i].YAdvance}, ShapeLightVertical={light[i].YAdvance}");
            }
        }

        [TestMethod]
        public void ShapeLightVertical_EmptyString_ReturnsEmptyArray()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeLightVertical(string.Empty);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Length);
        }

        #endregion


        #region Fallback tests (fonts without vmtx)

        [TestMethod]
        public void ShapeVertical_FontWithoutVmtx_FallsBackToHmtxAdvanceWidth()
        {
            // Arrange - Calibri has no vmtx table, fallback to hmtx should kick in
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "Roboto", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);
            Assert.IsNull(font.VmtxTable, "Roboto should not have a vmtx table");

            // Act
            var result = shaper.ShapeVertical("ABC");

            // Assert - YAdvance should still be positive (sourced from hmtx advanceWidth)
            Assert.AreEqual(3, result.Glyphs.Length);
            foreach (var glyph in result.Glyphs)
            {
                Assert.IsTrue(glyph.YAdvance > 0,
                    $"Glyph {glyph.GlyphId} has YAdvance {glyph.YAdvance}, expected > 0 via hmtx fallback");
                Assert.AreEqual((short)0, glyph.TopSideBearing,
                    "TopSideBearing should be 0 when falling back to hmtx");
            }
        }

        [TestMethod]
        public void ShapeVertical_FontWithoutVmtx_YAdvanceMatchesHmtxAdvanceWidth()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "Roboto", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeVertical("A");

            // Assert - YAdvance should equal hmtx advanceWidth for the same glyph
            var glyphId = result.Glyphs[0].GlyphId;
            var expectedAdvance = font.HmtxTable.GetAdvanceWidth(glyphId);
            Assert.AreEqual(expectedAdvance, result.Glyphs[0].YAdvance,
                "YAdvance via fallback should equal hmtx advanceWidth");
        }

        #endregion

        #region OriginalText and FontId tests

        [TestMethod]
        public void ShapeVertical_OriginalTextIsPreserved()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);
            var text = "日本語テスト";

            // Act
            var result = shaper.ShapeVertical(text);

            // Assert
            Assert.AreEqual(text, result.OriginalText);
        }

        [TestMethod]
        public void ShapeVertical_PrimaryFontGlyphs_HaveFontIdZero()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeVertical("日本語");

            // Assert - all glyphs should be from primary font (FontId == 0)
            foreach (var glyph in result.Glyphs)
            {
                Assert.AreEqual((byte)0, glyph.FontId,
                    $"Glyph {glyph.GlyphId} has FontId {glyph.FontId}, expected 0 for primary font");
            }
        }

        #endregion

        #region Surrogate pair tests

        [TestMethod]
        public void ShapeVertical_SurrogatePair_ProducesOneGlyphWithCharCountTwo()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);

            // U+20B9F 𠺟 - a CJK unified ideograph extension B character (surrogate pair in UTF-16)
            var text = "\uD842\uDF9F";
            Assert.AreEqual(2, text.Length, "Surrogate pair should be 2 UTF-16 code units");

            // Act
            var result = shaper.ShapeVertical(text);

            // Assert - surrogate pair should produce exactly one glyph
            Assert.AreEqual(1, result.Glyphs.Length,
                "A surrogate pair should produce exactly one glyph");
            Assert.AreEqual((byte)2, result.Glyphs[0].CharCount,
                "CharCount should be 2 for a surrogate pair");
            Assert.AreEqual((ushort)0, result.Glyphs[0].ClusterIndex,
                "ClusterIndex should point to the start of the surrogate pair");
        }

        [TestMethod]
        public void ShapeLightVertical_SurrogatePair_ProducesOneEntryWithCharCountTwo()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);
            var text = "\uD842\uDF9F";

            // Act
            var result = shaper.ShapeLightVertical(text);

            // Assert
            Assert.AreEqual(1, result.Length,
                "A surrogate pair should produce exactly one VerticalGlyphHeight");
            Assert.AreEqual((byte)2, result[0].CharCount,
                "CharCount should be 2 for a surrogate pair");
        }

        #endregion
    }
}