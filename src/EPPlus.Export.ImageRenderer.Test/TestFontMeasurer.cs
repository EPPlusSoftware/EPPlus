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
                   "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z " +
                   "Sooooo " +
                   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA " +
                   "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB" +
                   "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC";
            string fontName = "Aptos Narrow";
            double fontSize = 11.0d;
            //double MaxPixelWidth = 750d;
            double MaxPixelWidth = 750d;

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
            string actual = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
            string left=    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";

            Assert.AreEqual("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", strings[2]);
            string actual2 = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
            Assert.AreEqual("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", strings[3]);
            Assert.AreEqual("BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB", strings[4]);
            Assert.AreEqual("BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB", strings[5]);
            Assert.AreEqual("BBBBBBBBBCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC", strings[6]);
            Assert.AreEqual("CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC", strings[7]);
            Assert.AreEqual("CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC", strings[8]);
        }
    }
}
