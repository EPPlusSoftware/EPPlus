using EPPlus.Fonts.OpenType.TrueTypeMeasurer;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using Microsoft.VisualStudio.TestPlatform.CrossPlatEngine.Adapter;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class TextFragmentCollectionTests
    {
        [TestMethod]
        public void EnsureTextFragmentsAndWrapperWorkCorrectlyForLongParagraphs()
        {
            List<string> longString = new List<string> { "Hello World! a b c d e f g h i j k l m n o p q r s t u v w x y z \r\n" };
            List<float> fontsizes = new List<float> { 28f };

            var textFragments = new TextFragmentCollection(longString, fontsizes);

            var ttMeasurer = new FontMeasurerTrueType();
            List<MeasurementFont> fonts = new List<MeasurementFont>() { new MeasurementFont() { 
            FontFamily = "Aptos Narrow",
            Size = 28,
            Style = MeasurementFontStyles.Regular } };

            var outputLines = ttMeasurer.WrapMultipleTextFragments(textFragments, fonts, 225);

            Assert.AreEqual("Hello World! a b c d", outputLines[0]);
            Assert.AreEqual("e f g h i j k l m n o p q", outputLines[1]);
            Assert.AreEqual("r s t u v w x y z ", outputLines[2]);
        }

        //TODO: DOUBLE-CHECK BOLD+ITALIC for narrow later it seems innaccurate
        [TestMethod]
        public void MeasureBold()
        {
            var font2 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Bold
            };
            string test = "TextBox2";

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();
            var ttMeasurer = new FontMeasurerTrueType(font2);

            ttMeasurer.SetFont(font2);

            var widthInPoints = ttMeasurer.MeasureTextWidth(test);
            var inPixels = Math.Round(widthInPoints.PointToPixel(),0,MidpointRounding.AwayFromZero);

            Assert.AreEqual(54, inPixels);
        }

        [TestMethod]
        public void MeasureGoudy()
        {
            var font5 = new MeasurementFont()
            {
                FontFamily = "Goudy Stout",
                Size = 16,
                Style = MeasurementFontStyles.Regular
            };

            var text = "Goudy size";

            var ttMeasurer = new FontMeasurerTrueType(font5);

            var widthInPoints = ttMeasurer.MeasureTextWidth(text);
            var inPixels = Math.Round(widthInPoints.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            Assert.AreEqual(237, inPixels);
        }

        [TestMethod]
        public void MeasureWrappedWidths()
        {
            List<string> lstOfRichText = new() { "TextBox2", "ra underline", "La Strike", "Goudy size 16", "SvgSize 24" };

            var font2 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Bold
            };

            var font3 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Underline
            };

            var font4 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Strikeout
            };

            var font5 = new MeasurementFont()
            {
                FontFamily = "Goudy Stout",
                Size = 16,
                Style = MeasurementFontStyles.Regular
            };


            var font6 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 24,
                Style = MeasurementFontStyles.Regular
            };

            List<MeasurementFont> fonts = new() {font2, font3, font4, font5, font6 };

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();
            var ttMeasurer = new FontMeasurerTrueType(font2);

            var textFragments = new TextFragmentCollection(lstOfRichText);

            var wrappedLines = ttMeasurer.WrapMultipleTextFragmentsToTextLines(textFragments, fonts, maxSizePoints);

            var pixels1 = Math.Round(wrappedLines[0].RtFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels2 = Math.Round(wrappedLines[0].RtFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels3 = Math.Round(wrappedLines[0].RtFragments[2].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine = Math.Round(wrappedLines[0].Width.PointToPixel(),0, MidpointRounding.AwayFromZero);

            //~54 px
            Assert.AreEqual(54, pixels1);
            //~70 px aka 51.75pt
            Assert.AreEqual(70, pixels2);
            //~16-17 px This line contains a space at the end
            Assert.AreEqual(17, pixels3);

            //Total Width: ~140
            Assert.AreEqual(140d, pixelsWholeLine);

            var pixels21 = Math.Round(wrappedLines[1].RtFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels22 = Math.Round(wrappedLines[1].RtFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine2 = Math.Round(wrappedLines[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            //~34 px
            Assert.AreEqual(33, pixels21);
            // This line contains a space at the end
            Assert.AreEqual(248, pixels22);

            Assert.AreEqual(281, pixelsWholeLine2);

            var pixels31 = Math.Round(wrappedLines[2].RtFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels32 = Math.Round(wrappedLines[2].RtFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine3 = Math.Round(wrappedLines[2].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            Assert.AreEqual(35, pixels31);
            //This line does NOT contain a space at the end
            Assert.AreEqual(134, pixels32);

            
            Assert.AreEqual(169, pixelsWholeLine3);

            //Line 1 137 px 102.75 pt //result: 104.6328125 pt width "whole
            //Line 2 270 px 202.5 pt
            //Line 3 169 px 126.75 pt
        }
    }
}
