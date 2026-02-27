using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer;
using EPPlus.Fonts.OpenType.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using IntegrationTextFragment = EPPlus.Fonts.OpenType.Integration.TextFragment;
using TrueTypeTextFragment = EPPlus.Fonts.OpenType.TrueTypeMeasurer.TextFragment;

namespace EPPlus.Fonts.OpenType.Tests.Integration
{
    [TestClass]
    public class MeasurerComparisonTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region Single Text Measurement Comparison

        [TestMethod]
        public void Compare_MeasureSimpleText_ShouldBeClose()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            // Old measurer
            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);

            // New measurer
            var shaper = new TextShaper(font);
            var newMeasurer = new OpenTypeFontTextMeasurer(shaper);

            var measurementFont = new MeasurementFont
            {
                FontFamily = "Roboto",
                Size = 11
            };

            // Act
            var oldResult = oldMeasurer.MeasureText("Hello World", measurementFont);
            var newResult = newMeasurer.MeasureText("Hello World", measurementFont);

            // Assert
            Debug.WriteLine($"Old Width: {oldResult.Width}, New Width: {newResult.Width}");
            Debug.WriteLine($"Old Height: {oldResult.Height}, New Height: {newResult.Height}");
            Debug.WriteLine($"Old FontHeight: {oldResult.FontHeight}, New FontHeight: {newResult.FontHeight}");

            // Allow small tolerance for rounding differences
            double tolerance = 0.5; // points
            Assert.AreEqual(oldResult.Width, newResult.Width, tolerance,
                $"Width difference too large. Old: {oldResult.Width}, New: {newResult.Width}");
            Assert.AreEqual(oldResult.Height, newResult.Height, tolerance,
                $"Height difference too large. Old: {oldResult.Height}, New: {newResult.Height}");
        }

        [TestMethod]
        public void Compare_MeasureTextWithKerning_ShouldBeClose()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);

            var shaper = new TextShaper(font);
            var newMeasurer = new OpenTypeFontTextMeasurer(shaper);

            var measurementFont = new MeasurementFont
            {
                FontFamily = "Roboto",
                Size = 11
            };

            // Act - "AV" has kerning
            var oldResult = oldMeasurer.MeasureText("AV", measurementFont);
            var newResult = newMeasurer.MeasureText("AV", measurementFont);

            // Assert
            Debug.WriteLine($"Old Width: {oldResult.Width}, New Width: {newResult.Width}");

            double tolerance = 0.5;
            Assert.AreEqual(oldResult.Width, newResult.Width, tolerance,
                $"Kerned text width differs. Old: {oldResult.Width}, New: {newResult.Width}");
        }

        [TestMethod]
        public void Compare_MeasureMultiLineText_NewImplementationFixesBugs()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);
            oldMeasurer.MeasureWrappedTextCells = true;

            var shaper = new TextShaper(font);
            var newMeasurer = new OpenTypeFontTextMeasurer(shaper);
            newMeasurer.MeasureWrappedTextCells = true;

            var measurementFont = new MeasurementFont
            {
                FontFamily = "Roboto",
                Size = 11
            };

            string multiLineText = "Line 1\r\nLine 2\nLine 3";

            // Act
            var oldResult = oldMeasurer.MeasureText(multiLineText, measurementFont);
            var newResult = newMeasurer.MeasureText(multiLineText, measurementFont);

            // Measure each line individually for verification
            var lines = multiLineText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            float expectedMaxWidth = 0;
            foreach (var line in lines)
            {
                var lineResult = newMeasurer.MeasureText(line, measurementFont);
                expectedMaxWidth = Math.Max(expectedMaxWidth, lineResult.Width);
                Debug.WriteLine($"Line '{line}': {lineResult.Width}");
            }

            double expectedHeight = shaper.GetLineHeightInPoints(11.0f) * lines.Length;

            // Assert
            Debug.WriteLine("");
            Debug.WriteLine($"Expected max width (widest line): {expectedMaxWidth}");
            Debug.WriteLine($"Old Width: {oldResult.Width} (BUG: incorrect due to line break handling)");
            Debug.WriteLine($"New Width: {newResult.Width} (CORRECT: max of individual lines)");
            Debug.WriteLine("");
            Debug.WriteLine($"Expected total height ({lines.Length} lines): {expectedHeight}");
            Debug.WriteLine($"Old Height: {oldResult.Height} (BUG: returns single line height)");
            Debug.WriteLine($"New Height: {newResult.Height} (CORRECT: total height)");

            // Verify new implementation is correct
            Assert.AreEqual(expectedMaxWidth, newResult.Width, 0.1,
                "New measurer correctly returns max width of all lines");

            Assert.AreEqual(expectedHeight, newResult.Height, 0.1,
                "New measurer correctly returns total height");

            // Document that old implementation has bugs
            Debug.WriteLine("");
            Debug.WriteLine("DOCUMENTED BUGS in old FontMeasurerTrueType.MeasureText:");
            Debug.WriteLine("  1. Width calculation incorrect for multi-line text");
            Debug.WriteLine($"     - Old returns {oldResult.Width:F2} instead of correct {expectedMaxWidth:F2}");
            Debug.WriteLine("  2. Height returns single line instead of total");
            Debug.WriteLine($"     - Old returns {oldResult.Height:F2} instead of correct {expectedHeight:F2}");
            Debug.WriteLine("");
            Debug.WriteLine("✅ New OpenTypeFontTextMeasurer fixes both bugs");
        }

        #endregion

        #region Font Metrics Comparison

        [TestMethod]
        public void Compare_GetSingleLineSpacing_ShouldMatch()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var oldSpacing = oldMeasurer.GetSingleLineSpacing();
            var newSpacing = shaper.GetLineHeightInPoints(11.0f);

            // Assert
            Debug.WriteLine($"Old Spacing: {oldSpacing}, New Spacing: {newSpacing}");

            double tolerance = 0.1;
            Assert.AreEqual(oldSpacing, newSpacing, tolerance,
                $"Line spacing differs. Old: {oldSpacing}, New: {newSpacing}");
        }

        [TestMethod]
        public void Compare_GetBaseLine_ShouldMatch()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);

            // Act
            var oldBaseline = oldMeasurer.GetBaseLine();
            var newBaseline = shaper.GetBaseLineInPoints(11.0f);

            // Assert
            Debug.WriteLine($"Old Baseline: {oldBaseline}, New Baseline: {newBaseline}");

            double tolerance = 0.1;
            Assert.AreEqual(oldBaseline, newBaseline, tolerance,
                $"Baseline differs. Old: {oldBaseline}, New: {newBaseline}");
        }

        #endregion

        #region Text Wrapping Comparison

        [TestMethod]
        public void Compare_WrapSimpleText_ShouldGiveSameLines()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);

            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var measurementFont = new MeasurementFont
            {
                FontFamily = "Roboto",
                Size = 11
            };

            string text = "This is a long text that should wrap at some point";
            double maxWidth = 100; // pixels

            // Act
            var oldLines = oldMeasurer.MeasureAndWrapText(text, measurementFont, maxWidth);
            var newLines = layout.WrapText(text, 11f, maxWidth.PixelToPoint());

            // Assert
            Debug.WriteLine($"Old line count: {oldLines.Count}, New line count: {newLines.Count}");
            for (int i = 0; i < Math.Max(oldLines.Count, newLines.Count); i++)
            {
                var oldLine = i < oldLines.Count ? oldLines[i] : "(none)";
                var newLine = i < newLines.Count ? newLines[i] : "(none)";
                Debug.WriteLine($"Line {i}: Old='{oldLine}', New='{newLine}'");
            }

            Assert.AreEqual(oldLines.Count, newLines.Count,
                "Number of wrapped lines should match");

            for (int i = 0; i < oldLines.Count; i++)
            {
                Assert.AreEqual(oldLines[i], newLines[i],
                    $"Line {i} content differs. Old: '{oldLines[i]}', New: '{newLines[i]}'");
            }
        }

        [TestMethod]
        public void Compare_WrapTextWithPreExistingWidth_ShouldGiveSameLines()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var measurementFont = new MeasurementFont
            {
                FontFamily = "Roboto",
                Size = 11
            };

            string text = "continuation text that wraps";
            double maxWidth = 150; // pixels
            double preExistingWidth = 50; // pixels

            // Act
            var oldLines = oldMeasurer.MeasureAndWrapText(text, measurementFont, maxWidth, preExistingWidth);
            var newLines = layout.WrapText(text, 11f, maxWidth.PixelToPoint(), preExistingWidth.PixelToPoint());

            // Assert
            Debug.WriteLine($"Old line count: {oldLines.Count}, New line count: {newLines.Count}");
            for (int i = 0; i < Math.Max(oldLines.Count, newLines.Count); i++)
            {
                var oldLine = i < oldLines.Count ? oldLines[i] : "(none)";
                var newLine = i < newLines.Count ? newLines[i] : "(none)";
                Debug.WriteLine($"Line {i}: Old='{oldLine}', New='{newLine}'");
            }

            Assert.AreEqual(oldLines.Count, newLines.Count,
                "Number of wrapped lines should match with pre-existing width");
        }

        #endregion

        #region Rich Text Wrapping Comparison

        [TestMethod]
        public void Compare_WrapRichText_ShouldGiveSimilarLines()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper, FontFolders);

            var textFragments = new List<string> { "Hello ", "world ", "from ", "rich ", "text" };
            var fonts = new List<MeasurementFont>
            {
                new MeasurementFont { FontFamily = "Roboto", Size = 11 },
                new MeasurementFont { FontFamily = "Roboto", Size = 12, Style = MeasurementFontStyles.Bold },
                new MeasurementFont { FontFamily = "Roboto", Size = 11 },
                new MeasurementFont { FontFamily = "Roboto", Size = 11, Style = MeasurementFontStyles.Italic },
                new MeasurementFont { FontFamily = "Roboto", Size = 11 }
            };

            // New API uses IntegrationTextFragment
            var fragments = new List<IntegrationTextFragment>();
            for (int i = 0; i < textFragments.Count; i++)
            {
                fragments.Add(new IntegrationTextFragment
                {
                    Text = textFragments[i],
                    Font = fonts[i]
                });
            }

            double maxWidth = 100; // points

            // Act
            var oldLines = oldMeasurer.WrapMultipleTextFragments(textFragments, fonts, maxWidth);
            var newLines = layout.WrapRichText(fragments, maxWidth);

            // Assert
            Debug.WriteLine($"Old line count: {oldLines.Count}, New line count: {newLines.Count}");
            for (int i = 0; i < Math.Max(oldLines.Count, newLines.Count); i++)
            {
                var oldLine = i < oldLines.Count ? oldLines[i] : "(none)";
                var newLine = i < newLines.Count ? newLines[i] : "(none)";
                Debug.WriteLine($"Line {i}: Old='{oldLine}', New='{newLine}'");
            }

            // Note: May not match exactly due to improved shaping/kerning
            // But should be close
            Assert.IsTrue(Math.Abs(oldLines.Count - newLines.Count) <= 1,
                $"Line count should be similar. Old: {oldLines.Count}, New: {newLines.Count}");
        }

        #endregion

        #region Edge Cases Comparison

        [TestMethod]
        public void Compare_EmptyString_BothShouldReturnZero()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var newMeasurer = new OpenTypeFontTextMeasurer(shaper);

            var measurementFont = new MeasurementFont { FontFamily = "Roboto", Size = 11 };

            // Act
            var oldResult = oldMeasurer.MeasureText("", measurementFont);
            var newResult = newMeasurer.MeasureText("", measurementFont);

            // Assert
            Debug.WriteLine($"Old result: Width={oldResult.Width}, Height={oldResult.Height}");
            Debug.WriteLine($"New result: Width={newResult.Width}, Height={newResult.Height}");

            // Both should return 0 width for empty string
            Assert.AreEqual(0f, oldResult.Width, "Old measurer: Empty string should have 0 width");
            Assert.AreEqual(0f, newResult.Width, "New measurer: Empty string should have 0 width");

            // Note: Old measurer returns font height, new returns 0 (both reasonable)
            Debug.WriteLine($"Height difference: Old returns font height ({oldResult.Height}), New returns 0 ({newResult.Height})");
        }

        [TestMethod]
        public void Compare_SingleCharacter_ShouldMatch()
        {
            // Arrange
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

            var oldMeasurer = new FontMeasurerTrueType(11f, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var newMeasurer = new OpenTypeFontTextMeasurer(shaper);

            var measurementFont = new MeasurementFont { FontFamily = "Roboto", Size = 11 };

            // Act
            var oldResult = oldMeasurer.MeasureText("A", measurementFont);
            var newResult = newMeasurer.MeasureText("A", measurementFont);

            // Assert
            Debug.WriteLine($"Old Width: {oldResult.Width}, New Width: {newResult.Width}");

            double tolerance = 0.5;
            Assert.AreEqual(oldResult.Width, newResult.Width, tolerance);
        }

        #endregion
    }
}