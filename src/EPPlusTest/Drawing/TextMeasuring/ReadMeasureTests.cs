using EPPlusImageRenderer;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Drawing.Text;
using EPPlus.Fonts.OpenType.Utils;


namespace EPPlusTest.Drawing.TextMeasuring
{
    [TestClass]
    public class ReadMeasureTests: TestBase
    {
        [TestMethod]
        public void ReadShape()
        {
            using(var p = OpenTemplatePackage("ReadText.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var theShape = ws.Drawings[0].As.Shape;

                ws.Calculate();
                theShape.AdjustPositionAndSize();

                var width = theShape.Size.Width / 9525d;
                var height = theShape.Size.Height / 9525d;
                //var height = theShape.TextBodyItem.Paragraphs.GetSizeInPixels(theShape.GetPixelWidth(), theShape.GetPixelHeight(), theShape.Text, theShape.Font);
            }
        }
        [TestMethod]
        public void ReadLoremIpsum()
        {
            using (var p = OpenTemplatePackage("LoremIpsums20.xlsx"))
            {
                var ws1 = p.Workbook.Worksheets[0];
                var shape1 = ws1.Drawings[0].As.Shape;
                var someText = shape1.TextBody.Paragraphs;
            }
        }

        internal List<string> SplitIntoLines(string text)
        {
            return text.Split(new string[] { "\r\n" }, StringSplitOptions.None).ToList();
        }

        [TestMethod] 
        public void WrapMultipleFragments_SpacedEndWord()
        {
            List<string> txtRuns =
            [
                "H",
                "IJ",
                "K",
                "L",
                "M ",
                "NOPE",
            ];


            var mf = new MeasurementFont();
            mf.FontFamily = "Aptos Narrow";
            mf.Style = MeasurementFontStyles.Regular;
            mf.Size = 16;

            var mf2 = new MeasurementFont();
            mf2.FontFamily = "Goudy Stout";
            mf2.Style = MeasurementFontStyles.Regular;
            mf2.Size = 11;

            List<MeasurementFont> fonts =
            [
                mf,
                mf,
                mf,
                mf,
                mf
            ];

            fonts.Add(mf2);
            var engine = new OpenTypeFontEngine(cfg =>
            {
                cfg.SearchSystemDirectories = true;
            });
            var txtMeasurer = engine.GetTextLayoutEngineForFont(mf2);
            var maxWidth = 114d;

            var wrappedFragments = txtMeasurer.WrapRichText(txtRuns, fonts, maxWidth.PixelToPoint());

            Assert.AreEqual(2, wrappedFragments.Count);
            Assert.AreEqual("HIJKLM", wrappedFragments[0]);
            Assert.AreEqual("NOPE", wrappedFragments[1]);
        }

        [TestMethod]
        public void WrapMultipleFragments_LongPlusEndWord()
        {
            List<string> txtRuns =
            [
                "H",
                "IJ",
                "K",
                "L",
                "Mpqrstvdef",
                " ",
                "NOPE",
            ];


            var mf = new MeasurementFont();
            mf.FontFamily = "Aptos Narrow";
            mf.Style = MeasurementFontStyles.Regular;
            mf.Size = 16;

            var mf2 = new MeasurementFont();
            mf2.FontFamily = "Aptos Narrow";
            mf2.Style = MeasurementFontStyles.Regular;
            mf2.Size = 11;

            List<MeasurementFont> fonts =
            [
                mf,
                mf,
                mf,
                mf,
                mf
            ];

            fonts.Add(mf2);
            fonts.Add(mf2);

            var engine = new OpenTypeFontEngine(x => x.SearchSystemDirectories = true);
            var txtMeasurer = engine.GetTextLayoutEngineForFont(mf);

            var maxWidth = 114d;

            var wrappedFragments = txtMeasurer.WrapRichText(txtRuns, fonts, maxWidth.PixelToPoint());

            Assert.AreEqual(2, wrappedFragments.Count);
            Assert.AreEqual("HIJKLMpqrst", wrappedFragments[0]);
            Assert.AreEqual("vdef NOPE", wrappedFragments[1]);
        }

        [TestMethod]
        public void ReadRichTextBox()
        {
            using (var p = OpenTemplatePackage("paragraphBook.xlsx"))
            {
                var ws1 = p.Workbook.Worksheets[0];
                var shape1 = ws1.Drawings[0].As.Shape;
                var paragraphs = shape1.TextBody.Paragraphs;
                var someText = shape1.TextBody.Paragraphs.Text;
                var richText = shape1.RichText;

                shape1.GetSizeInPixels(out int width, out int height);

                var svg = shape1.ToSvg();

                var svgFile = GetOutputFile("", "paragraphBook.svg");

                //Create a file to write to.
                using (StreamWriter sw = svgFile.CreateText())
                {
                    sw.Write(svg);
                }
    
                SaveAndCleanup(p);
            }
        }
    }
}
