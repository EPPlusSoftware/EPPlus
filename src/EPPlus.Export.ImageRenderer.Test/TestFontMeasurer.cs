using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.RenderItems.Shared;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;

namespace TestProject1
{
    [TestClass]
    public class TestFontMeasurer
    {
        [TestMethod]
        public void CompareFontMeasurer3()
        {
            var fontName = "Aptos Narrow";
            var testStr = "Hello there⁴₂";
            var fontSize = 72.0d;

            FontMeasurerTrueType fontMeasurer = new FontMeasurerTrueType(fontSize, fontName);

            var exactWidth = fontMeasurer.MeasureTextWidthInPixels(testStr);
            var exactHeight = fontMeasurer.GetLargestPossibleHeightInPixels();

            //Ascent+Descent
            var lineSpacing = fontMeasurer.GetSingleLineSpacing().PointToPixel();
            
            //Distance between text baseline and top of box AKA Ascent
            var getBaseLine = fontMeasurer.GetBaseLine().PointToPixel();

            var approxHeight = fontMeasurer.InternalFontHeight().PointToPixel();

            var wholePixelWidth = TextUtils.RoundToWhole(exactWidth);
            var wholePixelHeight = TextUtils.RoundToWhole(getBaseLine);

            Assert.AreEqual(90d, wholePixelHeight, 0.1);
        }

        [TestMethod]
        public void TestWrapText()
        {
            string fontName = "Aptos Narrow";
            string testStr = "hello the most";
            double fontSize = 11.0d;
            double MaxPixelWidth = 52d;

            MeasurementFont mf = new MeasurementFont()
            {
                FontFamily = fontName,
                Size = (float)fontSize,
                Style = MeasurementFontStyles.Regular
            };


            FontMeasurerTrueType fontMeasurer = new FontMeasurerTrueType(fontSize, fontName);
            var strings = fontMeasurer.MeasureAndWrapText(testStr, mf, MaxPixelWidth);

            Assert.AreEqual("hello the", strings[0]);
            Assert.AreEqual("most", strings[1]);
        }

        [TestMethod]
        public void TestWrapTextLongContinous()
        {
            string testString = "Hello World! a b c d e f g h i j k l m n o p q r s t u v w x y z \r\n" +
                   "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z Sooooo " +
                   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA " +
                   "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB" +
                   "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBCCCCCC" +
                   "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC" +
                   "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC" +
                   "CCCCCCCCCCCCCCCCCCCCCCCCCC";
            double fontSize = 11.0d;
            //double MaxPixelWidth = 750d;
            double MaxPixelWidth = 750d;
            var fontName = "Aptos Narrow";
            MeasurementFont mf = new MeasurementFont()
            {
                FontFamily = fontName,
                Size = (float)fontSize,
                Style = MeasurementFontStyles.Regular
            };
            FontMeasurerTrueType fontMeasurer = new FontMeasurerTrueType(fontSize, fontName);
            var strings = fontMeasurer.MeasureAndWrapText(testString, mf, MaxPixelWidth);

            Assert.AreEqual("Hello World! a b c d e f g h i j k l m n o p q r s t u v w x y z ", strings[0]);
            Assert.AreEqual("A B C D E F G H I J K L M N O P Q R S T U V W X Y Z Sooooo", strings[1]);
            Assert.AreEqual("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", strings[2]);
            Assert.AreEqual("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", strings[3]);
            Assert.AreEqual("BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB", strings[4]);
            Assert.AreEqual("BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBCCCCCC", strings[5]);

            //Note: We wrap differently from Excel. We assume the kerning is applied correctly which means one extra 'C' fits in these two rows
            //And is therefore not part of the last one.
            Assert.AreEqual("CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC", strings[6]);
            Assert.AreEqual("CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC", strings[7]);
            Assert.AreEqual("CCCCCCCCCCCCCCCCCCCCCCCC", strings[8]);
        }
    }
}
