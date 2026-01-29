using EPPlus.Fonts.OpenType.TrueTypeMeasurer;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
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
    }
}
