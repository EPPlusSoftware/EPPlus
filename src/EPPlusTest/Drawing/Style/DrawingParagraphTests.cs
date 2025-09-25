using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text;

namespace EPPlusTest.Drawing.Style
{
    [TestClass]
    public class DrawingParagraphTests: TestBase
    {
        [TestMethod]
        public void EnsureExpectedParagraphCount()
        {
            using(var p = OpenPackage("DrawingParagraphCount.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("shapeParagraphs");
                var shape = ws.Drawings.AddShape("shape1", eShapeStyle.Sun);

                shape.Font.SetFromFont("Aptos Narrow", 11f);
                var font = shape.Font;
                font.Color = Color.Goldenrod;

                var paragraphs = shape.TextBody.Paragraphs;

                var para1 = paragraphs.Add("hello the most");

                var para2 = paragraphs.Add(" ");
                var para3 = paragraphs.Add("people");

                Assert.AreEqual(3, paragraphs.Count);

                SaveAndCleanup(p);
            }
        }
    }
}
