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
            var expectedWidth = 94d;
            var expectedHeight = 54d;

            var mf = new MeasurementFont() { FontFamily = "Aptos Narrow", Size = 11, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true);

            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }

        [TestMethod]
        public void TestNonStandardFontSizesMultiLine()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            var expectedWidth = 205d;
            var expectedHeight = 119d;

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
            var expectedWidth = 820d;
            var expectedHeight = 473d;

            var mf = new MeasurementFont() { FontFamily = "Aptos Narrow", Size = 96, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true, true);
            //0,072265625 * font size width correction for pixels
            //0,0442708333333333 * font size height correction for pixels
            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }
    }
}
