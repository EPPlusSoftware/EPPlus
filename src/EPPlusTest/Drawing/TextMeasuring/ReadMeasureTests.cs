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
    public class ReadMeasureTests : TestBase
    {
        [TestMethod]
        public void ReadShape()
        {
            using (var p = OpenTemplatePackage("ReadText.xlsx"))
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
