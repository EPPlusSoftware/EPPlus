/*************************************************************************************************
  Font Provider Unit Tests
  Tests for automatic emoji fallback functionality
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.FallbackFonts
{
    [TestClass]
    public class FontProviderTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private OpenTypeFont _robotoFont;

        [TestInitialize]
        public void TestSetup()
        {
            _robotoFont = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
        }

        [TestMethod]
        public void DefaultFontProvider_EmojiGlyph_ShouldUseFallbackFont()
        {
            // Arrange
            var shaper = new TextShaper(_robotoFont);

            // Act
            var shaped = shaper.Shape("😀");
            var usedFonts = shaper.GetUsedFonts().ToList();

            // Assert
            Assert.AreEqual(1, shaped.Glyphs.Length, "Should have 1 glyph");
            Assert.AreNotEqual((ushort)0, shaped.Glyphs[0].GlyphId, "Emoji should not be .notdef");
            Assert.AreEqual((byte)0, shaped.Glyphs[0].FontId, "Emoji is the only font used (FontId=0)");

            // Verify it's NOT the primary font
            Assert.AreEqual(1, usedFonts.Count, "Should only use one font (emoji fallback)");
            Assert.AreNotEqual(_robotoFont, usedFonts[0], "Should be emoji font, not Roboto");
        }

        [TestMethod]
        public void DefaultFontProvider_LatinText_ShouldUsePrimaryFont()
        {
            // Arrange
            var shaper = new TextShaper(_robotoFont);

            // Act
            var shaped = shaper.Shape("Hello World");

            // Assert
            foreach (var glyph in shaped.Glyphs)
            {
                Assert.AreEqual((byte)0, glyph.FontId, "All glyphs should be from primary font");
            }
        }

        [TestMethod]
        public void DefaultFontProvider_MixedTextAndEmoji_ShouldUseMultipleFonts()
        {
            // Arrange
            var shaper = new TextShaper(_robotoFont);

            // Act
            var shaped = shaper.Shape("Hello 😀 World");
            var usedFonts = shaper.GetUsedFonts().ToList();

            // Assert
            Assert.AreEqual(2, usedFonts.Count, "Should use 2 fonts (primary + emoji fallback)");
            Assert.AreEqual(_robotoFont, usedFonts[0], "First font should be primary");
            Assert.AreNotEqual(_robotoFont, usedFonts[1], "Second font should be emoji fallback");
        }

        [TestMethod]
        public void TextShaper_SurrogatePair_ShouldMapToSingleGlyph()
        {
            // Arrange
            var shaper = new TextShaper(_robotoFont);
            string text = "😀"; // U+1F600 = 2 chars in UTF-16

            // Act
            var shaped = shaper.Shape(text);

            // Assert
            Assert.AreEqual(2, text.Length, "Emoji should be 2 chars in .NET");
            Assert.AreEqual(1, shaped.Glyphs.Length, "Should map to 1 glyph");
            Assert.AreEqual((byte)2, shaped.Glyphs[0].CharCount, "Glyph should span 2 chars");
            Assert.AreEqual((ushort)0, shaped.Glyphs[0].ClusterIndex, "Should start at char 0");
        }

        [TestMethod]
        public void TextShaper_MultipleEmoji_ShouldMapCorrectly()
        {
            // Arrange
            var shaper = new TextShaper(_robotoFont);
            string text = "😀😁😂"; // 3 emoji = 6 chars in UTF-16

            // Act
            var shaped = shaper.Shape(text);

            // Assert
            Assert.AreEqual(6, text.Length, "3 emoji = 6 chars");
            Assert.AreEqual(3, shaped.Glyphs.Length, "Should map to 3 glyphs");

            Assert.AreEqual((ushort)0, shaped.Glyphs[0].ClusterIndex, "First emoji at char 0");
            Assert.AreEqual((byte)2, shaped.Glyphs[0].CharCount, "First emoji spans 2 chars");

            Assert.AreEqual((ushort)2, shaped.Glyphs[1].ClusterIndex, "Second emoji at char 2");
            Assert.AreEqual((byte)2, shaped.Glyphs[1].CharCount, "Second emoji spans 2 chars");

            Assert.AreEqual((ushort)4, shaped.Glyphs[2].ClusterIndex, "Third emoji at char 4");
            Assert.AreEqual((byte)2, shaped.Glyphs[2].CharCount, "Third emoji spans 2 chars");
        }
    }
}