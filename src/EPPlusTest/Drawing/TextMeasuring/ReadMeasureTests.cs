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
                //var height = theShape.TextBody.Paragraphs.GetSizeInPixels(theShape.GetPixelWidth(), theShape.GetPixelHeight(), theShape.Text, theShape.Font);
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
        public void ReadRichTextBox()
        {
            using (var p = OpenTemplatePackage("paragraphBookSimplified.xlsx"))
            {
                var ws1 = p.Workbook.Worksheets[0];
                var shape1 = ws1.Drawings[0].As.Shape;
                var paragraphs = shape1.TextBody.Paragraphs;
                var someText = shape1.TextBody.Paragraphs.Text;
                var richText = shape1.RichText;

                shape1.GetSizeInPixels(out int width, out int height);

                var txtMeasurer = new FontMeasurerTrueType();

                //List<string> textFragments = new List<string>();
                //List<MeasurementFont> fonts = new List<MeasurementFont>();

                //var lMargin = shape1.TextBody.LeftInsert.HasValue ? shape1.TextBody.LeftInsert.Value : 0;
                //var rMargin = shape1.TextBody.RightInsert.HasValue ? shape1.TextBody.RightInsert.Value : 0;

                //var maxWidth = width -
                //    lMargin - rMargin
                //    - paragraphs[2].LeftMargin - paragraphs[2].RightMargin;

                //foreach (var txtRun in paragraphs[2].TextRuns)
                //{
                //    textFragments.Add(txtRun.Text);
                //    fonts.Add(txtRun.GetMeasurementFont());
                //}

                //var wrappedFragments = txtMeasurer.WrapMultipleTextFragments(textFragments, fonts, maxWidth.PixelToPoint());

                //var lines = SplitIntoLines(someText);

                var lMargin = shape1.TextBody.LeftInsert.HasValue ? shape1.TextBody.LeftInsert.Value : 0;
                var rMargin = shape1.TextBody.RightInsert.HasValue ? shape1.TextBody.RightInsert.Value : 0;

                var wrappedStrings = paragraphs[2].TextRuns.MeasureAndWrapTextRuns(width -
                    lMargin - rMargin
                    - paragraphs[2].LeftMargin - paragraphs[2].RightMargin);

                var ir = new ImageRenderer();
                var svg = ir.RenderDrawingToSvg(shape1);

                var svgFile = GetOutputFile("", "paragraphBookSimplified.svg");

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
