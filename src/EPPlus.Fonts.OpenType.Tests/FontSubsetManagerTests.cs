using EPPlus.Fonts.OpenType;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class FontSubsetManagerTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        // Helper: Load a real font for testing
        private OpenTypeFont LoadTestFont()
        {
            // Adjust path to a font available in your test environment
            return TestFolderEngine.LoadFont("Roboto");
        }

        [TestMethod]
        public void CreateSubsettedProvider_WithAsciiText_ReturnsSubsettedPrimaryFont()
        {
            // Arrange
            var font = LoadTestFont();
            var manager = new FontSubsetManager(font);

            // Act
            manager.AddText("Hello World");
            var provider = manager.CreateSubsettedProvider();

            // Assert - The subset should be a different (smaller) font instance
            var subsetFont = provider.PrimaryFont;
            Assert.IsNotNull(subsetFont);
            Assert.IsTrue(subsetFont.IsSubset, "Primary font should be subsetted");

            // Verify the subset contains the glyphs we need
            foreach (char c in "Hello World")
            {
                ushort glyphId;
                Assert.IsTrue(
                    subsetFont.CmapTable.TryGetGlyphId(c, out glyphId),
                    $"Subset should contain glyph for '{c}'");
                Assert.AreNotEqual((ushort)0, glyphId, $"Glyph for '{c}' should not be .notdef");
            }
        }

        [TestMethod]
        public void CreateSubsettedProvider_WithEmoji_SubsetsFallbackFont()
        {
            // Arrange
            var font = LoadTestFont();
            var provider = new DefaultFontProvider(font);
            var manager = new FontSubsetManager(provider);

            // Act - Add text with emoji (U+1F600 = 😀, handled by Noto Emoji fallback)
            manager.AddText("Hello 😀");
            var subsettedProvider = manager.CreateSubsettedProvider();

            // Assert - Should have primary + at least one fallback
            var allFonts = subsettedProvider.GetAllFonts().ToList();
            Assert.IsTrue(allFonts.Count >= 2,
                "Should have primary font + emoji fallback font");

            // The fallback font should also be subsetted
            var fallbackFont = allFonts[1];
            Assert.IsTrue(fallbackFont.IsSubset,
                "Fallback (emoji) font should be subsetted");

            // The subsetted emoji font should be much smaller than the original
            var serialized = fallbackFont.Serialize();
            Assert.IsTrue(serialized.Length < 100 * 1024,
                $"Subsetted emoji font should be small, was {serialized.Length / 1024} KB");
        }

        [TestMethod]
        public void CreateSubsettedProvider_WithMultipleAddTextCalls_CollectsAllCodePoints()
        {
            // Arrange
            var font = LoadTestFont();
            var manager = new FontSubsetManager(font);

            // Act - Add text in multiple calls (simulates scanning multiple cells)
            manager.AddText("ABC");
            manager.AddText("DEF");
            manager.AddText("ADF"); // Overlapping characters
            var provider = manager.CreateSubsettedProvider();

            // Assert - All characters from all calls should be present
            var subsetFont = provider.PrimaryFont;
            foreach (char c in "ABCDEF")
            {
                ushort glyphId;
                Assert.IsTrue(
                    subsetFont.CmapTable.TryGetGlyphId(c, out glyphId),
                    $"Subset should contain glyph for '{c}'");
            }
        }

        [TestMethod]
        public void CreateSubsettedProvider_UnusedFallbackFontsAreExcluded()
        {
            // Arrange - DefaultFontProvider has Noto Emoji + Noto Math as fallbacks
            var font = LoadTestFont();
            var provider = new DefaultFontProvider(font);
            var manager = new FontSubsetManager(provider);

            // Act - Only ASCII text, no emoji or math symbols
            manager.AddText("Plain text only");
            var subsettedProvider = manager.CreateSubsettedProvider();

            // Assert - Should only have the primary font (no fallbacks needed)
            var allFonts = subsettedProvider.GetAllFonts().ToList();
            Assert.AreEqual(1, allFonts.Count,
                "Only primary font should be included when no fallback glyphs are used");
        }

        [TestMethod]
        public void AddText_WithNullOrEmpty_DoesNotThrow()
        {
            // Arrange
            var font = LoadTestFont();
            var manager = new FontSubsetManager(font);

            // Act & Assert - Should handle gracefully
            manager.AddText(null);
            manager.AddText("");
            manager.AddText("A"); // Then add real text

            var provider = manager.CreateSubsettedProvider();
            Assert.IsNotNull(provider.PrimaryFont);
        }
    }
}