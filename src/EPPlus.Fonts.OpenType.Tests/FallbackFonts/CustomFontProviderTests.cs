/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/24/2026         EPPlus Software AB           CustomFontProvider unit tests
 *************************************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.FallbackFonts
{
    [TestClass]
    public class CustomFontProviderTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private OpenTypeFont _robotoFont;
        private OpenTypeFont _notoEmojiFont;
        private OpenTypeFont _notoMathFont;

        // U+1F600 = 😀 (Grinning Face) - present in NotoEmoji, absent in Roboto
        private const uint EmojiCodePoint = 0x1F600;
        // U+0041 = 'A' - present in Roboto (and most text fonts)
        private const uint LatinA = 0x0041;
        // U+2A0C = ⨌ (Quadruple Integral) - present in Noto Sans Math, absent in Roboto and NotoEmoji
        private const uint IntegralCodePoint = 0x2A0C;
        // U+F0000 = Private Use Area - unlikely to be in any standard font
        private const uint PrivateUseCodePoint = 0xF0000;

        [TestInitialize]
        public void TestSetup()
        {
            _robotoFont = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);
            _notoEmojiFont = OpenTypeFonts.LoadFont("Noto Emoji", FontSubFamily.Regular);
            _notoMathFont = OpenTypeFonts.LoadFont("Noto Sans Math", FontSubFamily.Regular);
        }

        #region Constructor Tests

        [TestMethod]
        public void Constructor_WithValidFont_ShouldSetPrimaryFont()
        {
            // Act
            var provider = new CustomFontProvider(_robotoFont);

            // Assert
            Assert.AreEqual(_robotoFont, provider.PrimaryFont);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_WithNull_ShouldThrowArgumentNullException()
        {
            // Act
            new CustomFontProvider(null);
        }

        #endregion

        #region AddFallback Tests

        [TestMethod]
        public void AddFallback_WithValidFont_ShouldBeIncludedInGetAllFonts()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);

            // Act
            provider.AddFallback(_notoEmojiFont);
            var allFonts = provider.GetAllFonts().ToList();

            // Assert
            Assert.AreEqual(2, allFonts.Count, "Should have primary + 1 fallback");
            Assert.AreEqual(_robotoFont, allFonts[0], "First should be primary font");
            Assert.AreEqual(_notoEmojiFont, allFonts[1], "Second should be fallback font");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void AddFallback_WithNull_ShouldThrowArgumentNullException()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);

            // Act
            provider.AddFallback(null);
        }

        [TestMethod]
        public void AddFallback_MultipleFonts_ShouldPreserveOrder()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);

            // Act
            provider.AddFallback(_notoEmojiFont);
            provider.AddFallback(_robotoFont); // Same font again, to test ordering
            var allFonts = provider.GetAllFonts().ToList();

            // Assert
            Assert.AreEqual(3, allFonts.Count, "Should have primary + 2 fallbacks");
            Assert.AreEqual(_robotoFont, allFonts[0], "First should be primary");
            Assert.AreEqual(_notoEmojiFont, allFonts[1], "Second should be first added fallback");
            Assert.AreEqual(_robotoFont, allFonts[2], "Third should be second added fallback");
        }

        #endregion

        #region TryGetGlyphFont - Primary Font Tests

        [TestMethod]
        public void TryGetGlyphFont_LatinChar_ShouldUsePrimaryFont()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_notoEmojiFont);

            // Act
            bool found = provider.TryGetGlyphFont(LatinA, out var font, out var glyphId);

            // Assert
            Assert.IsTrue(found, "Should find glyph for 'A'");
            Assert.AreEqual(_robotoFont, font, "Should use primary font for Latin chars");
            Assert.AreNotEqual((ushort)0, glyphId, "GlyphId should not be .notdef");
        }

        [TestMethod]
        public void TryGetGlyphFont_CharInBothFonts_ShouldPreferPrimaryFont()
        {
            // Arrange - both Roboto and NotoEmoji might have basic chars,
            // but primary should always be preferred
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_notoEmojiFont);

            // Act
            bool found = provider.TryGetGlyphFont(LatinA, out var font, out var glyphId);

            // Assert
            Assert.IsTrue(found);
            Assert.AreEqual(_robotoFont, font, "Primary font should always be preferred when it has the glyph");
        }

        #endregion

        #region TryGetGlyphFont - Fallback Tests

        [TestMethod]
        public void TryGetGlyphFont_EmojiNotInPrimary_ShouldUseFallbackFont()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_notoEmojiFont);

            // Verify preconditions: Roboto should NOT have the emoji glyph
            bool robotoHasEmoji = _robotoFont.CmapTable.TryGetGlyphId(EmojiCodePoint, out _);
            Assert.IsFalse(robotoHasEmoji, "Precondition: Roboto should not contain emoji U+1F600");

            // Verify preconditions: NotoEmoji SHOULD have the emoji glyph
            bool notoHasEmoji = _notoEmojiFont.CmapTable.TryGetGlyphId(EmojiCodePoint, out _);
            Assert.IsTrue(notoHasEmoji, "Precondition: NotoEmoji should contain emoji U+1F600");

            // Act
            bool found = provider.TryGetGlyphFont(EmojiCodePoint, out var font, out var glyphId);

            // Assert
            Assert.IsTrue(found, "Should find emoji glyph in fallback font");
            Assert.AreEqual(_notoEmojiFont, font, "Should return NotoEmoji as the fallback font");
            Assert.AreNotEqual((ushort)0, glyphId, "GlyphId should not be .notdef");
        }

        [TestMethod]
        public void TryGetGlyphFont_EmojiWithMultipleFallbacks_ShouldUseFirstMatchingFallback()
        {
            // Arrange - add Roboto (no emoji) first, then NotoEmoji (has emoji)
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_robotoFont);      // First fallback: also no emoji
            provider.AddFallback(_notoEmojiFont);    // Second fallback: has emoji

            // Act
            bool found = provider.TryGetGlyphFont(EmojiCodePoint, out var font, out var glyphId);

            // Assert
            Assert.IsTrue(found, "Should find emoji in second fallback");
            Assert.AreEqual(_notoEmojiFont, font, "Should return NotoEmoji (second fallback), not Roboto (first fallback)");
            Assert.AreNotEqual((ushort)0, glyphId);
        }

        [TestMethod]
        public void TryGetGlyphFont_MathSymbol_ShouldUseThirdFallback()
        {
            // Arrange - three-level fallback chain: Roboto → NotoEmoji → NotoSansMath
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_notoEmojiFont);
            provider.AddFallback(_notoMathFont);

            // Verify preconditions
            bool robotoHasMath = _robotoFont.CmapTable.TryGetGlyphId(IntegralCodePoint, out _);
            Assert.IsFalse(robotoHasMath, "Precondition: Roboto should not contain ∫ (U+222B)");

            bool emojiHasMath = _notoEmojiFont.CmapTable.TryGetGlyphId(IntegralCodePoint, out _);
            Assert.IsFalse(emojiHasMath, "Precondition: NotoEmoji should not contain ∫ (U+222B)");

            bool mathHasMath = _notoMathFont.CmapTable.TryGetGlyphId(IntegralCodePoint, out _);
            Assert.IsTrue(mathHasMath, "Precondition: Noto Sans Math should contain ∫ (U+222B)");

            // Act
            bool found = provider.TryGetGlyphFont(IntegralCodePoint, out var font, out var glyphId);

            // Assert
            Assert.IsTrue(found, "Should find ∫ in third fallback (Noto Sans Math)");
            Assert.AreEqual(_notoMathFont, font, "Should return Noto Sans Math as the resolving font");
            Assert.AreNotEqual((ushort)0, glyphId, "GlyphId should not be .notdef");
        }

        [TestMethod]
        public void TryGetGlyphFont_ThreeFallbacks_EachFontResolvesItsOwnDomain()
        {
            // Arrange - full chain: Roboto → NotoEmoji → NotoSansMath
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_notoEmojiFont);
            provider.AddFallback(_notoMathFont);

            // Act & Assert - Latin 'A' should resolve from primary (Roboto)
            bool foundLatin = provider.TryGetGlyphFont(LatinA, out var latinFont, out var latinGlyphId);
            Assert.IsTrue(foundLatin);
            Assert.AreEqual(_robotoFont, latinFont, "Latin 'A' should come from Roboto (primary)");
            Assert.AreNotEqual((ushort)0, latinGlyphId);

            // Act & Assert - Emoji should resolve from first fallback (NotoEmoji)
            bool foundEmoji = provider.TryGetGlyphFont(EmojiCodePoint, out var emojiFont, out var emojiGlyphId);
            Assert.IsTrue(foundEmoji);
            Assert.AreEqual(_notoEmojiFont, emojiFont, "Emoji should come from NotoEmoji (1st fallback)");
            Assert.AreNotEqual((ushort)0, emojiGlyphId);

            // Act & Assert - Math symbol should resolve from second fallback (NotoSansMath)
            bool foundMath = provider.TryGetGlyphFont(IntegralCodePoint, out var mathFont, out var mathGlyphId);
            Assert.IsTrue(foundMath);
            Assert.AreEqual(_notoMathFont, mathFont, "∫ should come from Noto Sans Math (2nd fallback)");
            Assert.AreNotEqual((ushort)0, mathGlyphId);

            // Act & Assert - Unknown code point should fail gracefully
            bool foundUnknown = provider.TryGetGlyphFont(PrivateUseCodePoint, out var unknownFont, out var unknownGlyphId);
            Assert.IsFalse(foundUnknown);
            Assert.AreEqual(_robotoFont, unknownFont, "Not found should return primary font");
            Assert.AreEqual((ushort)0, unknownGlyphId);
        }

        #endregion

        #region TryGetGlyphFont - Not Found Tests

        [TestMethod]
        public void TryGetGlyphFont_CharNotInAnyFont_ShouldReturnFalseWithPrimaryFont()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_notoEmojiFont);

            // Act
            bool found = provider.TryGetGlyphFont(PrivateUseCodePoint, out var font, out var glyphId);

            // Assert
            Assert.IsFalse(found, "Should not find glyph in Private Use Area");
            Assert.AreEqual(_robotoFont, font, "Should return primary font even when not found");
            Assert.AreEqual((ushort)0, glyphId, "GlyphId should be .notdef (0)");
        }

        [TestMethod]
        public void TryGetGlyphFont_NoFallbacks_CharNotInPrimary_ShouldReturnFalse()
        {
            // Arrange - no fallbacks added at all
            var provider = new CustomFontProvider(_robotoFont);

            // Act
            bool found = provider.TryGetGlyphFont(EmojiCodePoint, out var font, out var glyphId);

            // Assert
            Assert.IsFalse(found, "Should not find emoji when no fallbacks are configured");
            Assert.AreEqual(_robotoFont, font, "Should still return primary font");
            Assert.AreEqual((ushort)0, glyphId, "Should return .notdef");
        }

        #endregion

        #region GetAllFonts Tests

        [TestMethod]
        public void GetAllFonts_NoFallbacks_ShouldReturnOnlyPrimary()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);

            // Act
            var allFonts = provider.GetAllFonts().ToList();

            // Assert
            Assert.AreEqual(1, allFonts.Count, "Should only contain primary font");
            Assert.AreEqual(_robotoFont, allFonts[0]);
        }

        [TestMethod]
        public void GetAllFonts_WithFallbacks_ShouldReturnPrimaryFirstThenFallbacks()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_notoEmojiFont);

            // Act
            var allFonts = provider.GetAllFonts().ToList();

            // Assert
            Assert.AreEqual(2, allFonts.Count);
            Assert.AreEqual(_robotoFont, allFonts[0], "Primary font should be first");
            Assert.AreEqual(_notoEmojiFont, allFonts[1], "Fallback should come after primary");
        }

        [TestMethod]
        public void GetAllFonts_CalledMultipleTimes_ShouldReturnConsistentResults()
        {
            // Arrange
            var provider = new CustomFontProvider(_robotoFont);
            provider.AddFallback(_notoEmojiFont);

            // Act
            var firstCall = provider.GetAllFonts().ToList();
            var secondCall = provider.GetAllFonts().ToList();

            // Assert
            Assert.AreEqual(firstCall.Count, secondCall.Count);
            for (int i = 0; i < firstCall.Count; i++)
            {
                Assert.AreEqual(firstCall[i], secondCall[i], $"Font at index {i} should be the same across calls");
            }
        }

        #endregion
    }
}