using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.DataHolders
{
    [TestClass]
    public class TextLineSimpleTests
    {
        [TestMethod]
        public void TestLineFragmentAbstraction()
        {

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();

            var fragments = GetTextFragments();

            var layout = OpenTypeFonts.GetTextLayoutEngineForFont(fragments[0].Font);
            var wrappedLines = layout.WrapRichTextLines(fragments, maxSizePoints);

            var line1 = wrappedLines[0];
        }

        List<TextFragment> GetTextFragments()
        {
            List<string> lstOfRichText = new() { "TextBox\r\na\r\n", "TextBox2", "ra underline", "La Strike", "Goudy size 16", "SvgSize 24" };

            var font1 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Regular
            };

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
            var fragments = new List<TextFragment>();

            for (int i = 0; i < lstOfRichText.Count(); i++)
            {
                var currentFrag = new TextFragment() { Text = lstOfRichText[i], Font = fonts[i] };
                fragments.Add(currentFrag);
            }

            return fragments;
        }
    }
}
