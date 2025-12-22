/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/21/2025         EPPlus Software AB           Basic subsetting tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class BasicSubsettingTests : FontTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            FontDirectoriesTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void Subset_Abc_RoundtripValidation()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            // Save for inspection
            SaveFont("subset_Roboto_abc.ttf", parsedFont);

            // Assert: Check table presence
            Assert.AreEqual(11, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("head"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("name"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("maxp"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hhea"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hmtx"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("loca"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("glyf"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("cmap"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("post"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("OS/2"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("GSUB"));

            // Validate all tables
            FontTestHelper.AssertFontValid(parsedFont);

            // ✅ FIXED: abc should have NO ligatures
            int ligatureCount = FontTestHelper.CountLigatures(parsedFont);
            Assert.AreEqual(0, ligatureCount, "abc should have NO ligatures");

            // Verify glyph count (approximately)
            int expectedGlyphs = 3;      // a, b, c
            expectedGlyphs += 1;         // + space (U+0020)
            expectedGlyphs += 1;         // + .notdef (GID 0)
            expectedGlyphs += 5;         // + variants from Single Substitution

            Assert.AreEqual((ushort)expectedGlyphs, parsedFont.MaxpTable.numGlyphs);
            Assert.AreEqual((ushort)expectedGlyphs, parsedFont.HheaTable.numberOfHMetrics);

            // Verify space exists in cmap
            Assert.IsTrue(parsedFont.CmapTable.ContainsChar(32));
        }

        [TestMethod]
        public void Subset_Fiffig_WithFullValidation()
        {
            // Arrange
            var fontName = "Roboto";
            var font = OpenTypeFonts.GetFontData(FontFolders, fontName, FontSubFamily.Regular, true);
            var subsetFont = font.CreateSubset("fiffig");

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            // Save for inspection
            SaveFont("subset_Roboto_fiffig.ttf", parsedFont);

            // Assert
            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_SingleChar_ShouldWork()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Mulish", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a' });

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            SaveFont("subset_Mulish_a.ttf", parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_MultipleChars_ShouldWork()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] {
                'F', 'l', 'y', 'g', 'a', 'n', 'd', 'e', 'b', 'ä', 'c', 'k', 's', 'i', 'r', 'ö', 'h', 'w', 'p', 'å',
                'm', 'j', 'u', 't', 'v', 'o', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0'
            });

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            SaveFont("subset_Roboto_flygande_bäckasiner.ttf", parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_RoundtripHelper_ShouldWork()
        {
            // Using FontTestHelper.RoundtripSubset
            var parsedFont = FontTestHelper.RoundtripSubset("Roboto", "test", FontFolders);

            SaveFont("subset_Roboto_test_via_helper.ttf", parsedFont);

            Assert.IsNotNull(parsedFont);
            Assert.IsTrue(parsedFont.MaxpTable.numGlyphs > 0);
        }
    }
}