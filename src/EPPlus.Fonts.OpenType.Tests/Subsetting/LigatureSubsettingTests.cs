/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/21/2025         EPPlus Software AB           Ligature subsetting tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tests.Helpers;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    /// <summary>
    /// Tests for GSUB ligature subsetting functionality
    /// </summary>
    [TestClass]
    public class LigatureSubsettingTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void Subset_Abc_ShouldHaveNoLigatures()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset(new[] { 'a', 'b', 'c' });

            SaveFontForCurrentTest(subset);

            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.AreEqual(0, ligCount, "abc should have NO ligatures");
        }

        [TestMethod]
        public void Subset_Fiffig_ShouldHaveThreeLigatures()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("fiffig");

            SaveFontForCurrentTest(subset);

            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.AreEqual(3, ligCount, "fiffig should have fi, ff, ffi ligatures");
        }

        [TestMethod]
        public void Subset_Fi_ShouldHaveFiLigature()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("fi");

            SaveFontForCurrentTest(subset);

            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.IsTrue(ligCount >= 1, "fi should have at least fi ligature");
        }

        [TestMethod]
        public void Subset_Ff_ShouldHaveFfLigature()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("ff");

            SaveFontForCurrentTest(subset);

            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.IsTrue(ligCount >= 1, "ff should have ff ligature");
        }

        [TestMethod]
        [DataRow("fi")]
        [DataRow("ffi")]
        [DataRow("ff")]
        [DataRow("fl")]
        public void Subset_CommonLigatures_ShouldWork(string text)
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset(text);

            SaveFontForCurrentTest(subset, text);

            Assert.IsNotNull(subset);
            FontTestHelper.AssertFontValid(subset);

            // Should have at least one ligature
            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.IsTrue(ligCount > 0, $"{text} should have ligatures");
        }

        [TestMethod]
        public void Subset_OnlyF_ShouldNotCrash()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("f");

            SaveFontForCurrentTest(subset);

            // Should not crash, may or may not have ligatures
            Assert.IsNotNull(subset);
            Assert.IsTrue(subset.MaxpTable.numGlyphs > 0);
            FontTestHelper.AssertFontValid(subset);

        }

        [TestMethod]
        public void Subset_Office_ShouldHaveFfiLigature()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("office");

            SaveFontForCurrentTest(subset);

            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.IsTrue(ligCount >= 1, "office should trigger ffi ligature");
        }

        [TestMethod]
        public void Subset_HasLigatureLookupType()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("fi");

            bool hasLigatures = FontTestHelper.HasGsubLookupType(subset, 4);
            Assert.IsTrue(hasLigatures, "Should have ligature lookup (Type 4)");
        }
    }
}