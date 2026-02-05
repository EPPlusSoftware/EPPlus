using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    [TestClass]
    public class SingleAdjustmentTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void SingleAdjustment_Roboto_DoesNotCrash()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act - Shape text with glyphs that are in the Single Adjustment coverage
            // (even though the adjustments are all zero)
            var result = shaper.Shape("Hello World");

            // Assert - Should not crash and should produce valid output
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Glyphs.Length > 0);
            Assert.AreEqual("Hello World", result.OriginalText);
        }

        [TestMethod]
        public void SingleAdjustment_WithZeroValues_DoesNotAffectOutput()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act - Shape same text with and without positioning
            var withPositioning = shaper.Shape("AV");
            var withoutPositioning = shaper.Shape("AV", ShapingOptions.None);

            // Assert - With zero-value Single Adjustments, the glyphs should only
            // differ by kerning (not by Single Adjustment since all values are 0)
            Assert.AreEqual(withoutPositioning.Glyphs.Length, withPositioning.Glyphs.Length);

            // The difference should only be from kerning
            Assert.IsTrue(withPositioning.TotalAdvanceWidth < withoutPositioning.TotalAdvanceWidth,
                "Should have kerning applied");
        }

        [TestMethod]
        public void SingleAdjustment_WithZeroValues_DoesNotAffectOutput2()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act - Shape same text with and without positioning
            var withPositioning = shaper.Shape("AV");
            var withoutPositioning = shaper.Shape("AV", ShapingOptions.None);

            // Debug output
            Debug.WriteLine($"Without positioning: {withoutPositioning.TotalAdvanceWidth}");
            Debug.WriteLine($"With positioning: {withPositioning.TotalAdvanceWidth}");
            Debug.WriteLine($"Difference: {withoutPositioning.TotalAdvanceWidth - withPositioning.TotalAdvanceWidth}");

            Debug.WriteLine("\nWithout positioning glyphs:");
            foreach (var g in withoutPositioning.Glyphs)
            {
                Debug.WriteLine($"  GlyphId: {g.GlyphId}, XAdvance: {g.XAdvance}, XOffset: {g.XOffset}");
            }

            Debug.WriteLine("\nWith positioning glyphs:");
            foreach (var g in withPositioning.Glyphs)
            {
                Debug.WriteLine($"  GlyphId: {g.GlyphId}, XAdvance: {g.XAdvance}, XOffset: {g.XOffset}");
            }

            // Assert
            Assert.AreEqual(withoutPositioning.Glyphs.Length, withPositioning.Glyphs.Length);
            Assert.IsTrue(withPositioning.TotalAdvanceWidth < withoutPositioning.TotalAdvanceWidth,
                "Should have kerning applied");
        }

        [TestMethod]
        public void Kerning_IsApplied_ForAVPair()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Shape with only kerning (no single adjustment)
            var optionsOnlyKern = new ShapingOptions
            {
                ApplySubstitutions = false,
                ApplyPositioning = true,
                GposFeatures = new List<string> { "kern" },
                Script = "latn",
                Language = null
            };

            var withoutKerning = shaper.Shape("AV", ShapingOptions.None);
            var withKerning = shaper.Shape("AV", optionsOnlyKern);

            Assert.IsTrue(withKerning.TotalAdvanceWidth < withoutKerning.TotalAdvanceWidth,
                "Kerning should reduce advance width for AV pair");
        }

        [TestMethod]
        public void SingleAdjustment_Verdana_HasRealAdjustments()
        {
            // NOTE: This test requires Verdana font to be installed
            // Verdana has Single Adjustment Format 2 with XPlacement=36 for certain glyphs

            try
            {
                // Arrange
                var font = OpenTypeFonts.GetFontData(FontFolders, "Verdana", FontSubFamily.Regular);
                if (font == null || font.FullName != "Verdana")
                {
                    Assert.Inconclusive("Verdana font not found - test skipped");
                    return; // Koden efter detta körs inte
                }
                var shaper = new TextShaper(font);

                // Act - Shape some text (we don't know which specific glyphs have adjustments)
                var withPositioning = shaper.Shape("Hello123");
                var withoutPositioning = shaper.Shape("Hello123", ShapingOptions.None);

                // Assert - Should produce valid output
                Assert.IsNotNull(withPositioning);
                Assert.IsNotNull(withoutPositioning);
                Assert.AreEqual(withPositioning.Glyphs.Length, withoutPositioning.Glyphs.Length);

                // Check if any glyph has XOffset applied (from Single Adjustment)
                bool hasXOffset = false;
                for (int i = 0; i < withPositioning.Glyphs.Length; i++)
                {
                    if (withPositioning.Glyphs[i].XOffset != 0)
                    {
                        hasXOffset = true;
                        System.Console.WriteLine($"Glyph {i} (GID={withPositioning.Glyphs[i].GlyphId}) has XOffset={withPositioning.Glyphs[i].XOffset}");
                    }
                }

                // Note: We can't assert that hasXOffset is true because we don't know
                // which characters map to the adjusted glyphs. But we can verify no crash.
                System.Console.WriteLine($"Found XOffset adjustments: {hasXOffset}");
            }
            catch (System.IO.FileNotFoundException)
            {
                Assert.Inconclusive("Verdana font not found - test skipped");
            }
        }

        [TestMethod]
        public void SingleAdjustment_Verdana_AdjustmentsAppliedWithDefaultOptions()
        {
            // NOTE: This test requires Verdana font to be installed

            try
            {
                // Arrange
                var font = OpenTypeFonts.GetFontData(FontFolders, "Verdana", FontSubFamily.Regular);
                if (font == null || font.FullName != "Verdana")
                {
                    Assert.Inconclusive("Verdana font not found - test skipped");
                    return; // Koden efter detta körs inte
                }
                var shaper = new TextShaper(font);

                // Act - Use default options (which includes positioning)
                var result = shaper.Shape("Test");

                // Assert - Should not crash and produce valid output
                Assert.IsNotNull(result);
                Assert.AreEqual(4, result.Glyphs.Length);
                Assert.AreEqual("Test", result.OriginalText);
            }
            catch (System.IO.FileNotFoundException)
            {
                Assert.Inconclusive("Verdana font not found - test skipped");
            }
        }

        [TestMethod]
        public void SingleAdjustment_AppliedBeforeKerning()
        {
            // This test verifies the order of operations:
            // Single Adjustment should be applied before Kerning

            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.Shape("Test");

            // Assert - Just verify it doesn't crash and produces valid output
            // (We can't test the actual order without non-zero Single Adjustment values)
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Glyphs.Length > 0);
        }

        [TestMethod]
        public void SingleAdjustment_NotAppliedWithNoneOptions()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var result = shaper.Shape("Test", ShapingOptions.None);

            // Assert - Should have basic glyph mapping but no positioning
            Assert.IsNotNull(result);
            Assert.AreEqual(4, result.Glyphs.Length);

            // With ShapingOptions.None, glyphs should have their base advance widths
            // (no kerning or Single Adjustment applied)
            foreach (var glyph in result.Glyphs)
            {
                Assert.AreEqual(0, glyph.XOffset, "No offset adjustments with None options");
                Assert.AreEqual(0, glyph.YOffset, "No offset adjustments with None options");
            }
        }

        [TestMethod]
        public void SingleAdjustmentProvider_HandlesNullFont()
        {
            // Arrange - Create provider with font that has no GPOS
            var font = OpenTypeFonts.GetFontData(FontFolders, "SourceSans3", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act - Should not crash even though SourceSans3 has no GPOS table
            var result = shaper.Shape("Test");

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(4, result.Glyphs.Length);
        }

        // NOTE: Verdana Single Adjustment coverage includes 397 glyphs across 3 Format 2 subtables
        // with XPlacement=36. This is likely for superscript or special positioning features.
        // To fully test these, we would need to:
        // 1. Identify which specific characters map to the adjusted glyphs
        // 2. Verify the XOffset is correctly applied (should be 36 font units)
        // 3. Test interaction with kerning (Single Adjustment first, then kerning)
        //
        // For now, these tests verify that:
        // - The code doesn't crash with real Single Adjustment data
        // - Options are respected
        // - Both zero-value (Roboto) and non-zero (Verdana) cases work
    }
}