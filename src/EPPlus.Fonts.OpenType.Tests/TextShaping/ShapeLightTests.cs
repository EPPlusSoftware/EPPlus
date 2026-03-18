/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/18/2026         EPPlus Software AB           ShapeLight multi-font tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    [TestClass]
    public class ShapeLightTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void ShapeLight_SimpleText_ReturnsSameGlyphCountAsShape()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);

            // Act
            var full = shaper.Shape("Hello");
            shaper.ResetFontTracking();
            var light = shaper.ShapeLight("Hello");

            // Assert
            Assert.AreEqual(full.Glyphs.Length, light.Glyphs.Length,
                "ShapeLight should produce same number of glyphs as Shape");
        }

        [TestMethod]
        public void ShapeLight_SimpleText_HasFontUnitsPerEm()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeLight("Hello");

            // Assert
            Assert.IsNotNull(result.FontUnitsPerEm, "Should have FontUnitsPerEm");
            Assert.AreEqual(1, result.FontUnitsPerEm.Length, "Single font should have 1 entry");
            Assert.AreEqual(font.HeadTable.UnitsPerEm, result.FontUnitsPerEm[0]);
        }

        [TestMethod]
        public void ShapeLight_EmojiOnly_UsesFallbackFont()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeLight("😀😁😂");

            // Assert
            Assert.AreEqual(3, result.Glyphs.Length, "Should have 3 glyphs for 3 emojis");
            Assert.IsNotNull(result.FontUnitsPerEm);
            Assert.IsTrue(result.FontUnitsPerEm.Length >= 1, "Should have at least one font");

            // All glyphs should have FontId 0 (the only used font is emoji fallback)
            foreach (var glyph in result.Glyphs)
            {
                Assert.AreEqual(0, glyph.FontId, "All emoji glyphs should be FontId 0");
                Assert.IsTrue(glyph.XAdvance > 0, "Emoji glyphs should have positive width");
            }
        }

        [TestMethod]
        public void ShapeLight_MixedTextAndEmoji_HasMultipleFonts()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeLight("Hi 😀 there");

            // Assert
            Assert.IsNotNull(result.FontUnitsPerEm);
            Assert.AreEqual(2, result.FontUnitsPerEm.Length,
                "Should have 2 fonts (primary + emoji fallback)");

            // Verify emoji glyph has different FontId than text glyphs
            var textFontId = result.Glyphs[0].FontId;
            var emojiFontId = result.Glyphs.First(g => g.ClusterIndex == 3).FontId; // '😀' starts at index 3
            Assert.AreNotEqual(textFontId, emojiFontId,
                "Emoji should use different font than text");
        }

        [TestMethod]
        public void ShapeLight_GetWidthInPoints_ConsistentWithShape()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);
            float fontSize = 12f;

            // Act
            var full = shaper.Shape("Hello World");
            float fullWidth = full.GetWidthInPoints(fontSize);

            shaper.ResetFontTracking();
            var light = shaper.ShapeLight("Hello World");
            float lightWidth = light.GetWidthInPoints(fontSize);

            // Assert — ShapeLight uses simplified kerning so allow small difference
            Assert.AreEqual(fullWidth, lightWidth, fullWidth * 0.05f,
                "ShapeLight width should be within 5% of Shape width");
        }

        [TestMethod]
        public void ShapeLight_FillCharWidths_ProducesCorrectWidths()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);
            string text = "ABC";
            float fontSize = 12f;
            var charWidths = new double[text.Length];

            // Act
            var result = shaper.ShapeLight(text);
            result.FillCharWidths(fontSize, charWidths, text.Length);

            // Assert
            for (int i = 0; i < text.Length; i++)
            {
                Assert.IsTrue(charWidths[i] > 0,
                    $"Character '{text[i]}' at index {i} should have positive width");
            }

            // Total of char widths should match GetWidthInPoints
            double totalCharWidths = charWidths.Sum();
            float shapedWidth = result.GetWidthInPoints(fontSize);
            Assert.AreEqual(shapedWidth, totalCharWidths, 0.01,
                "Sum of char widths should match total shaped width");
        }

        [TestMethod]
        public void ShapeLight_FillCharWidths_MixedEmoji_CorrectPerGlyphScale()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);
            string text = "A😀B";
            float fontSize = 12f;
            var charWidths = new double[text.Length];

            // Act
            var result = shaper.ShapeLight(text);
            result.FillCharWidths(fontSize, charWidths, text.Length);

            // Assert
            Assert.IsTrue(charWidths[0] > 0, "'A' should have positive width");
            // charWidths[1] is high surrogate of emoji — should have the emoji width
            Assert.IsTrue(charWidths[1] > 0, "Emoji (at surrogate position) should have positive width");
            // charWidths[2] is low surrogate — typically 0 (width is on the high surrogate)
            // charWidths[3] is 'B'
            Assert.IsTrue(charWidths[text.Length - 1] > 0, "'B' should have positive width");

            // Total should match
            double totalCharWidths = charWidths.Sum();
            float shapedWidth = result.GetWidthInPoints(fontSize);
            Assert.AreEqual(shapedWidth, totalCharWidths, 0.01,
                "Sum of char widths should match total shaped width");
        }

        [TestMethod]
        public void ShapeLight_EmptyString_ReturnsEmptyResult()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.ShapeLight("");

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Glyphs.Length);
            Assert.IsNotNull(result.FontUnitsPerEm);
            Assert.AreEqual(0f, result.GetWidthInPoints(12f));
        }

        [TestMethod]
        public void ShapeLight_TextEmojiAndMath_UsesThreeFonts()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto");
            var shaper = new TextShaper(font);

            // U+1D400 = 𝐀 (Mathematical Bold Capital A) — not in Roboto or Noto Emoji,
            // should fall back to Noto Sans Math.
            // If U+1D400 isn't covered, try U+2200 (∀) or U+222B (∫).
            string mathChar = "\u2200";
            string text = "Hi😀" + mathChar;
            float fontSize = 12f;

            // Act
            var result = shaper.ShapeLight(text);

            // Assert — three distinct fonts
            Assert.IsNotNull(result.FontUnitsPerEm);
            Assert.AreEqual(3, result.FontUnitsPerEm.Length,
                $"Expected 3 fonts (primary + emoji + math), got {result.FontUnitsPerEm.Length}");

            // Verify three distinct FontIds are present
            var distinctFontIds = result.Glyphs.Select(g => g.FontId).Distinct().OrderBy(id => id).ToArray();
            Assert.AreEqual(3, distinctFontIds.Length,
                $"Expected 3 distinct FontIds, got [{string.Join(", ", distinctFontIds)}]");

            // All glyphs should have valid (non-zero) advance widths
            foreach (var glyph in result.Glyphs)
            {
                Assert.IsTrue(glyph.XAdvance > 0,
                    $"Glyph at ClusterIndex {glyph.ClusterIndex} (FontId={glyph.FontId}) should have positive width");
            }

            // FillCharWidths should still sum correctly
            var charWidths = new double[text.Length];
            result.FillCharWidths(fontSize, charWidths, text.Length);
            double totalCharWidths = charWidths.Sum();
            float shapedWidth = result.GetWidthInPoints(fontSize);
            Assert.AreEqual(shapedWidth, totalCharWidths, 0.01,
                "Sum of char widths should match total shaped width across all three fonts");
        }
    }
}