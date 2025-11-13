using EPPlus.Export.ImageRenderer.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Tests
{
    [TestClass]
    public class TestTextContainer
    {
        [TestMethod]
        public void TestDynamicSizeSingleLine()
        {
            string content = "TextBox";
            var expectedWidth = 48d;
            var expectedHeight = 18d;

            var mf = new MeasurementFont() { FontFamily = "Aptos", Size = 11, Style = MeasurementFontStyles.Regular };
            
            var container = new TextContainer(content, mf, true);

            Assert.AreEqual(expectedWidth, Math.Round(container.Width,0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height,0));
        }

        [TestMethod]
        public void TestDynamicSizeMultipleLines()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            var expectedWidth = 93d;
            var expectedHeight = 54d;

            //0.57 inches
            //0.98 inches

            var mf = new MeasurementFont() { FontFamily = "Aptos Narrow", Size = 11, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true);

            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }

        [TestMethod]
        public void TestNonStandardFontSizesMultiLine()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            var expectedWidth = 203d;
            var expectedHeight = 117d;

            var mf = new MeasurementFont() { FontFamily = "Aptos Narrow", Size = 24, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true, true);
            //0,056640625  * font size width correction for pixels (minimum)
            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }

        [TestMethod]
        public void TestNonStandardFontSizesMultiLineLargeFont()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            var expectedWidth = 813d;
            var expectedHeight = 469d;

            var mf = new MeasurementFont() { FontFamily = "Aptos Narrow", Size = 96, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true, true);
            //0,072265625 * font size width correction for pixels
            //0,0442708333333333 * font size height correction for pixels
            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }

        [TestMethod]
        public void TestNonStandardFontSizesMultiLineLargeFontGoudyStout()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            //This is the actual height in excel.
            //var expectedHeight = 533d;
            //It is unclear why as the Max possible glyph height in pixels for each line is
            //175.125 pixels. *3 = 525.375
            //Best guess is it adds 4 pixels per row for potential "outline"

            //This is the glyph yMax and YMin
            var expectedHeight = 525d;
            //Excel has 2332d but they add some kind of internal buffer/minX spacing for lineEnding
            var expectedWidth = 2309d;

            var mf = new MeasurementFont() { FontFamily = "Goudy Stout", Size = 96, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true, true);
            //0,072265625 * font size width correction for pixels
            //0,0442708333333333 * font size height correction for pixels
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }
    }
}
