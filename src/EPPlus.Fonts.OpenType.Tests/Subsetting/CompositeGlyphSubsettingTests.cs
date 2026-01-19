/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/21/2025         EPPlus Software AB           Composite glyph subsetting tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class CompositeGlyphSubsettingTests : FontTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            FontDirectoriesTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void Subset_Roboto_With_ÅÄÖ_Should_Work()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            // Get the original å
            var ågId = font.CmapTable.MapCharToGlyph('å');
            var åglyph = font.GlyfTable.GetGlyph((ushort)ågId);

            var subset = font.CreateSubset("Testar åäö ÅÄÖ och även é û č ć đ ł".Distinct());

            // Save for inspection (CI/CD safe)
            SaveFont("Roboto-subset-aao.ttf", subset);

            // Verify that 'å' actually has a composite glyph
            var åGlyphId = subset.CmapTable.MapCharToGlyph('å');
            var glyph = subset.GlyfTable.GetGlyph((ushort)åGlyphId);

            Assert.IsTrue(glyph.Header.numberOfContours < 0, "å should be composite");
            Assert.IsTrue(glyph.CompositeData.Components.Count > 0);
        }

        [TestMethod]
        public void Subset_Mulish_With_ÅÄÖ_Should_Work()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Mulish", FontSubFamily.Regular);
            var subset = font.CreateSubset("Testar åäö ÅÄÖ och även é û č ć đ ł".Distinct());

            // Save for inspection (CI/CD safe)
            SaveFont("Mulish-subset-aao.ttf", subset);

            // Verify that 'å' actually has a composite glyph
            var åGlyphId = subset.CmapTable.MapCharToGlyph('å');
            var glyph = subset.GlyfTable.GetGlyph((ushort)åGlyphId);

            Assert.IsTrue(glyph.Header.numberOfContours < 0, "å should be composite");
            Assert.IsTrue(glyph.CompositeData.Components.Count > 0);
        }

        [TestMethod]
        public void Subset_BIZUDGothic_With_ÅÄÖ_Should_Work()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "BIZUDGothic", FontSubFamily.Regular);
            var subset = font.CreateSubset("Testar åäö ÅÄÖ och även é û č ć đ ł".Distinct());

            // Save for inspection (CI/CD safe)
            SaveFont("BIZUDGothic-subset-aao.ttf", subset);

            // Verify that 'å' is a simple glyph in BIZUDGothic (not composite)
            var åGlyphId = subset.CmapTable.MapCharToGlyph('å');
            var glyph = subset.GlyfTable.GetGlyph((ushort)åGlyphId);

            Assert.AreEqual(4, glyph.Header.numberOfContours, "BIZUDGothic å should be simple with 4 contours");
            Assert.IsNotNull(glyph.SimpleData);
            Assert.IsNull(glyph.CompositeData);
        }
    }
}