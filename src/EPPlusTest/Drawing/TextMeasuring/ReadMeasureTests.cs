using EPPlusImageRenderer;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;


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

                //List<ExcelParagraphTextRunBase> runs = new();
                //foreach (var paragraph in paragraphs)
                //{
                //    foreach(var textRun in paragraph.TextRuns)
                //    {
                //        runs.Add(textRun);
                //    }
                //}

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
