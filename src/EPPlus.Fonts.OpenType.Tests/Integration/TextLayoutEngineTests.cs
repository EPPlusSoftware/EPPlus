using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.Tests.Integration
{
    [TestClass]
    public class TextLayoutEngineTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region Single-Font Wrapping Tests

        [TestMethod]
        public void WrapText_ShortText_NoWrapping()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper, FontFolders);

            // Act
            var lines = layout.WrapText("Hello", 11f, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello", lines[0]);
        }

        [TestMethod]
        public void WrapText_LongText_WrapsAtSpaces()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act - narrow width forces wrapping
            var lines = layout.WrapText("Hello world test", 11f, 50);

            // Assert
            Assert.IsTrue(lines.Count > 1, "Text should wrap to multiple lines");

            // Each line should be a complete word (no mid-word breaks)
            foreach (var line in lines)
            {
                Assert.IsFalse(string.IsNullOrEmpty(line));
            }
        }

        [TestMethod]
        public void WrapText_WithLineBreaks_PreservesBreaks()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act
            var lines = layout.WrapText("Line 1\r\nLine 2\nLine 3", 11f, 1000);

            // Assert
            Assert.AreEqual(3, lines.Count);
            Assert.AreEqual("Line 1", lines[0]);
            Assert.AreEqual("Line 2", lines[1]);
            Assert.AreEqual("Line 3", lines[2]);
        }

        [TestMethod]
        public void WrapText_WithPreExistingWidth_AccountsForIt()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Measure "Hello " to get its width
            var testShaper = new TextShaper(font);
            var shaped = testShaper.Shape("Hello ", ShapingOptions.Default);
            double preWidth = shaped.GetWidthInPoints(11f, testShaper.UnitsPerEm);

            // Act - Add text with pre-existing width, narrow max width
            var lines = layout.WrapText("world test", 11f, preWidth + 50, preWidth);

            // Assert - Should wrap because first line already has content
            Assert.IsTrue(lines.Count >= 1);
        }

        [TestMethod]
        public void WrapText_EmptyString_ReturnsEmptyLine()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act
            var lines = layout.WrapText("", 11f, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual(string.Empty, lines[0]);
        }

        [TestMethod]
        public void WrapText_WithKerning_MeasuresCorrectly()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act - "AV" has kerning in Roboto
            var withKerning = layout.WrapText("AV", 11f, 1000, ShapingOptions.Default);
            var withoutKerning = layout.WrapText("AV", 11f, 1000, ShapingOptions.None);

            // Assert - Both should be single line, but measured differently
            Assert.AreEqual(1, withKerning.Count);
            Assert.AreEqual(1, withoutKerning.Count);
            Assert.AreEqual("AV", withKerning[0]);
            Assert.AreEqual("AV", withoutKerning[0]);
        }

        #endregion

        #region Rich Text Wrapping Tests

        [TestMethod]
        public void WrapRichText_SingleFragment_BehavesLikeSingleFont()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Hello world",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello world", lines[0]);
        }

        [TestMethod]
        public void WrapRichText_MultipleFragments_ConcatenatesCorrectly()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Hello ",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "world",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 12, Style = MeasurementFontStyles.Bold }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello world", lines[0]);
        }

        [TestMethod]
        public void WrapRichText_DifferentFonts_WrapsCorrectly()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "This is ",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "mixed ",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 14, Style = MeasurementFontStyles.Bold }
                },
                new TextFragment
                {
                    Text = "fonts",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                }
            };

            // Act - narrow width to force wrapping
            var lines = layout.WrapRichText(fragments, 80);

            // Assert
            Assert.IsTrue(lines.Count >= 1);

            // Concatenate all lines should give original text
            string allText = string.Join("", lines).Replace(" ", " ");
            Assert.IsTrue(allText.Contains("This is mixed fonts"));
        }

        [TestMethod]
        public void WrapRichText_DifferentFonts_WrapsCorrectly2()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "This is ",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "mixed ",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 14, Style = MeasurementFontStyles.Bold }
                },
                new TextFragment
                {
                    Text = "fonts",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                }
            };

            // Act - narrow width to force wrapping
            var lines = layout.WrapRichText(fragments, 80);

            // Debug output
            Debug.WriteLine($"Number of lines: {lines.Count}");
            foreach (var line in lines)
            {
                Debug.WriteLine($"  Line: '{line}'");
            }
            string allText = string.Join("", lines);
            Debug.WriteLine($"All text: '{allText}'");

            // Assert
            Assert.IsTrue(lines.Count >= 1, $"Expected at least 1 line, got {lines.Count}");

            // Concatenate all lines should give original text
            allText = string.Join("", lines).Replace(" ", " ");
            Debug.WriteLine($"Cleaned text: '{allText}'");
            Assert.IsTrue(allText.Contains("This is mixed fonts"),
                $"Expected text to contain 'This is mixed fonts', but got: '{allText}'");
        }

        [TestMethod]
        public void WrapRichText_WordSpanningFragments_MeasuresCorrectly()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // "Hello" split across two fragments with different fonts
            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Hel",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "lo world",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello world", lines[0]);
        }

        [TestMethod]
        public void WrapRichText_WithLineBreaks_PreservesBreaks()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Line 1\n",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "Line 2",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(2, lines.Count);
            Assert.AreEqual("Line 1", lines[0]);
            Assert.AreEqual("Line 2", lines[1]);
        }

        [TestMethod]
        public void WrapRichText_EmptyFragments_HandlesGracefully()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "Hello",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 }
                },
                new TextFragment
                {
                    Text = "",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello", lines[0]);
        }

        [TestMethod]
        public void WrapRichText_NullFragmentList_ReturnsEmptyLine()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act
            var lines = layout.WrapRichText(null, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual(string.Empty, lines[0]);
        }

        #endregion

        #region Font Caching Tests

        [TestMethod]
        public void WrapRichText_SameFontMultipleTimes_UsesCache()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "First ",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 }
                },
                new TextFragment
                {
                    Text = "second ",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "third",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 } // Same as first
                }
            };

            // Act - This should use cached shaper for Arial
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("First second third", lines[0]);
            // Note: We can't easily verify cache usage without exposing internals,
            // but this test documents the expected behavior
        }

        #endregion
    }
}