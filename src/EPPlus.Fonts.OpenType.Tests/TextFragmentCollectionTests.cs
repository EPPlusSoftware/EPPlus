using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class TextFragmentCollectionTests
    {
        [TestMethod]
        public void EnsureTextFragmentsAndWrapperWorkCorrectlyForLongParagraphs()
        {
            var shaper = OpenTypeFonts.GetTextShaper("Aptos Narrow", FontSubFamily.Regular);
            var layout = new TextLayoutEngine(shaper);

            var outputLines = layout.WrapText("Hello World! a b c d e f g h i j k l m n o p q r s t u v w x y z \r\n", 28f, 225);

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

            var measurer = OpenTypeFonts.GetShaperForFont(font2);
            var widthInPoints = measurer.GetDescentInPoints(font2.Size);

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

            var measurer = OpenTypeFonts.GetShaperForFont(font5);
            var widthInPoints = measurer.GetDescentInPoints(font5.Size);

            var inPixels = Math.Round(widthInPoints.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            Assert.AreEqual(237, inPixels);
        }

        [TestMethod]
        public void MeasureWrappedWidthsWithInternalLineBreaks()
        {
            List<string> lstOfRichText = new() { "TextBox\r\na\r\n", "TextBox2", "ra underline", "La Strike", "Goudy size 16", "SvgSize 24" };

            var font1 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Regular
            }; ;

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

            List<MeasurementFont> fonts = new() { font1, font2, font3, font4, font5, font6 };

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();
            var ttMeasurer = OpenTypeFonts.GetTextLayoutEngineForFont(font1);

            var textFragments = new TextFragmentCollection(lstOfRichText);

            var wrappedLines = ttMeasurer.WrapRichTextLines(fonts, maxSizePoints);

            var line1 = wrappedLines[0];

            var pixels11 = Math.Round(line1.RtFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeline1 = Math.Round(line1.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            Assert.AreEqual(44, pixels11);
            Assert.AreEqual(pixels11, pixelsWholeline1);

            var line2 = wrappedLines[1];

            var pixels21 = Math.Round(line2.RtFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeline2 = Math.Round(line2.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            Assert.AreEqual(7, pixels21);
            Assert.AreEqual(pixels21, pixelsWholeline2);

            var line3 = wrappedLines[2];

            var pixels31 = Math.Round(line3.RtFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels32 = Math.Round(line3.RtFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels33 = Math.Round(line3.RtFragments[2].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine3 = Math.Round(line3.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            //~54 px
            Assert.AreEqual(54, pixels31);
            //~70 px aka 51.75pt
            Assert.AreEqual(70, pixels32);
            //~16-17 px This line contains a space at the end
            Assert.AreEqual(17, pixels33);

            //Total Width: ~140
            Assert.AreEqual(140d, pixelsWholeLine3);

            var line4 = wrappedLines[3];

            var pixels41 = Math.Round(line4.RtFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels42 = Math.Round(line4.RtFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine4 = Math.Round(line4.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            //~34 px
            Assert.AreEqual(33, pixels41);
            // This line contains a space at the end
            Assert.AreEqual(248, pixels42);

            Assert.AreEqual(281, pixelsWholeLine4);

            var line5 = wrappedLines[4];

            var pixels51 = Math.Round(line5.RtFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels52 = Math.Round(line5.RtFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine5 = Math.Round(line5.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            Assert.AreEqual(35, pixels51);
            //This line does NOT contain a space at the end
            Assert.AreEqual(134, pixels52);


            Assert.AreEqual(169, pixelsWholeLine5);
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
        }

        [TestMethod]
        public void CorrectTextLinesAreReturnedWhenSmallMaxWidth()
        {
            var defaultFont = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Regular
            };

            var ttMeasurer = new FontMeasurerTrueType(defaultFont);

            var txt = "This is my text";

            //This should be small enough to put each word on a new row.
            var maxWidth = 21.5d;

            var txtLines = ttMeasurer.MeasureAndWrapTextLines(txt, defaultFont, maxWidth);

            Assert.AreEqual(4, txtLines.Count);
            Assert.AreEqual("This", txtLines[0].Text);
            Assert.AreEqual("is", txtLines[1].Text);
            Assert.AreEqual("my", txtLines[2].Text);
            Assert.AreEqual("text", txtLines[3].Text);
        }
    }
}
