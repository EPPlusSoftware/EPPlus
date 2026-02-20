 /*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           Vertical subsetting tests (vhea/vmtx)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tests.Helpers;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class VerticalSubsettingTests : FontTestBase
    {
        public override TestContext TestContext { get; set; }

        [ClassInitialize]
        public static void ClassInitialize(TestContext ctx)
        {
            FontDirectoriesTestHelper.ClassInitialize(ctx);
        }

        #region vhea/vmtx presence tests

        [TestMethod]
        public void Subset_CjkFont_SubsetContainsVheaTable()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            Assert.IsNotNull(font.VheaTable, "BIZ UDGothic should have a vhea table");

            // Act
            var subset = font.CreateSubset("日本語");
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            SaveFontForCurrentTest(parsed);

            // Assert
            Assert.IsNotNull(parsed.VheaTable, "Subset should contain vhea table");
        }

        [TestMethod]
        public void Subset_CjkFont_SubsetContainsVmtxTable()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            Assert.IsNotNull(font.VmtxTable, "BIZ UDGothic should have a vmtx table");

            // Act
            var subset = font.CreateSubset("日本語");
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            SaveFontForCurrentTest(parsed);

            // Assert
            Assert.IsNotNull(parsed.VmtxTable, "Subset should contain vmtx table");
        }

        [TestMethod]
        public void Subset_FontWithoutVmtx_SubsetDoesNotContainVmtxTable()
        {
            // Arrange - Roboto has no vmtx/vhea
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "Roboto", FontSubFamily.Regular, true);
            Assert.IsNull(font.VmtxTable, "Roboto should not have a vmtx table");

            // Act
            var subset = font.CreateSubset("ABC");
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            // Assert - vmtx should not be introduced by subsetting
            Assert.IsNull(parsed.VmtxTable, "Subset of font without vmtx should not contain vmtx table");
        }

        #endregion

        #region vhea correctness tests

        [TestMethod]
        public void Subset_CjkFont_VheaNumberOfVMetricsMatchesGlyphCount()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);

            // Act
            var subset = font.CreateSubset("日本語");
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            SaveFontForCurrentTest(parsed);

            // Assert - NumberOfVMetrics must equal numGlyphs (same simplification as hmtx)
            Assert.AreEqual(
                parsed.MaxpTable.numGlyphs,
                parsed.VheaTable.NumberOfVMetrics,
                "vhea.NumberOfVMetrics should equal numGlyphs in subset");
        }

        #endregion

        #region vmtx correctness tests

        [TestMethod]
        public void Subset_CjkFont_VmtxAdvanceHeightPreservedForSubsettedGlyphs()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);
            var text = "日";

            // Get original glyph ID and advance height before subsetting
            ushort originalGlyphId;
            font.CmapTable.TryGetGlyphId('日', out originalGlyphId);
            var originalAdvanceHeight = font.VmtxTable.GetAdvanceHeight(originalGlyphId);

            // Act
            var subset = font.CreateSubset(text);
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            SaveFontForCurrentTest(parsed);

            // Assert - resolve new glyph ID in subset and verify advance height is preserved
            ushort subsetGlyphId;
            parsed.CmapTable.TryGetGlyphId('日', out subsetGlyphId);
            var subsetAdvanceHeight = parsed.VmtxTable.GetAdvanceHeight(subsetGlyphId);

            Assert.AreEqual(originalAdvanceHeight, subsetAdvanceHeight,
                $"AdvanceHeight for '日' should be preserved after subsetting " +
                $"(original={originalAdvanceHeight}, subset={subsetAdvanceHeight})");
        }

        [TestMethod]
        public void Subset_CjkFont_VmtxEntryCountMatchesGlyphCount()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);

            // Act
            var subset = font.CreateSubset("東京");
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            SaveFontForCurrentTest(parsed);

            // Assert - VMetrics count must equal numGlyphs
            Assert.AreEqual(
                parsed.MaxpTable.numGlyphs,
                parsed.VmtxTable.VMetrics.Count,
                "vmtx.VMetrics.Count should equal numGlyphs in subset");
        }

        [TestMethod]
        public void Subset_CjkFont_PassesValidationAfterSubsetting()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontDataOpen(FontFolders, "BIZ UDGothic", FontSubFamily.Regular, true);

            // Act
            var subset = font.CreateSubset("日本語テスト");
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            SaveFontForCurrentTest(parsed);

            // Assert - full font validation should pass without errors
            FontTestHelper.AssertFontValid(parsed, FontValidationSeverity.Warning);
        }

        #endregion
    }
}