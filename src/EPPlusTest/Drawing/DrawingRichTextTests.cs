using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class DrawingRichTextTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DrawingRichText.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Richtext");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);

            File.Copy(fileName, dirName + "\\DrawingRichTextRead.xlsx", true);
        }

        [TestMethod]
        public void AddThreeParagraphsAndValidate()
        {
            var shape = _ws.Drawings.AddShape("shape1", eShapeStyle.Rect);
            shape.RichText.Add("Line1");
            var r2 = shape.RichText.Add("L", true);
            r2.Fill.Style = eFillStyle.SolidFill;
            r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent2);
            r2 = shape.RichText.Add("i");
            r2.Fill.Style = eFillStyle.SolidFill;
            r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent3);
            r2 = shape.RichText.Add("n");
            r2.Fill.Style = eFillStyle.SolidFill;
            r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent4);
            r2 = shape.RichText.Add("e");
            r2.Fill.Style = eFillStyle.SolidFill;
            r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent5);
            r2 = shape.RichText.Add("2");
            r2.Fill.Style = eFillStyle.SolidFill;
            r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent6);


            var r3 = shape.RichText.Add("Line3", true);
            r3.Bold = true;
            r3.Italic = true;
            r3.LatinFont = "Times New Roman";
            r3.Size = 19.5F;

            Assert.AreEqual("Line1\r\nLine2\r\nLine3", shape.Text);
            Assert.AreEqual("Line1\r\nLine2\r\nLine3", shape.RichText.Text);

            Assert.AreEqual(7, shape.RichText.Count);
            Assert.IsTrue(shape.RichText[0].IsFirstInParagraph);
            Assert.IsTrue(shape.RichText[0].IsLastInParagraph);
            Assert.IsTrue(shape.RichText[1].IsFirstInParagraph);
            Assert.IsFalse(shape.RichText[1].IsLastInParagraph);
            Assert.IsFalse(shape.RichText[2].IsFirstInParagraph);
            Assert.IsFalse(shape.RichText[2].IsLastInParagraph);
            Assert.IsFalse(shape.RichText[3].IsFirstInParagraph);
            Assert.IsFalse(shape.RichText[3].IsLastInParagraph);
            Assert.IsFalse(shape.RichText[4].IsFirstInParagraph);
            Assert.IsFalse(shape.RichText[4].IsLastInParagraph);
            Assert.IsFalse(shape.RichText[5].IsFirstInParagraph);
            Assert.IsTrue(shape.RichText[5].IsLastInParagraph);
            Assert.IsTrue(shape.RichText[6].IsFirstInParagraph);
            Assert.IsTrue(shape.RichText[6].IsLastInParagraph);
        }
        [TestMethod]
        public void ReadThreeParagraphsAndValidate()
        {
            AssertIfNotExists("DrawingRichTextReadFunctional.xlsx");
            using (var p = OpenPackage("DrawingRichTextReadFunctional.xlsx"))
            {
                var shape = (ExcelShape)p.Workbook.Worksheets[0].Drawings["shape1"];
                Assert.AreEqual("Line1\r\nLine2\r\nLine3", shape.Text);
                Assert.AreEqual("Line1\r\nLine2\r\nLine3", shape.RichText.Text);

                Assert.AreEqual(7, shape.RichText.Count);
                Assert.IsTrue(shape.RichText[0].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[0].IsLastInParagraph);
                Assert.IsTrue(shape.RichText[1].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[1].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[2].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[2].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[3].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[3].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[4].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[4].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[5].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[5].IsLastInParagraph);
                Assert.IsTrue(shape.RichText[6].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[6].IsLastInParagraph);
            }
        }
        [TestMethod]
        public void AddEmptyParagraphFirst()
        {
            var shape = _ws.Drawings.AddShape("shape2", eShapeStyle.Rect);
            shape.SetPosition(20, 0, 0, 0);
            shape.RichText.Add("", true);
            shape.RichText.Add("SecondLine", true);
            var r2 = shape.RichText.Add("    ", true);
            r2.UnderLine = OfficeOpenXml.Style.eUnderLineType.Single;
            Assert.AreEqual(3, shape.RichText.Count);
            Assert.AreEqual("", shape.RichText[0].Text);
            Assert.AreEqual("SecondLine", shape.RichText[1].Text);
            Assert.AreEqual("    ", shape.RichText[2].Text);
        }
        [TestMethod]
        public void ReadParagraphsShapes()
        {
            using (var p = OpenTemplatePackage("Paragraphs.xlsx"))
            {
                var ws1 = p.Workbook.Worksheets[0];
                var pg1 = ws1.Drawings[0].As.Shape.TextBody.Paragraphs;

                Assert.AreEqual(7, pg1.Count);
                Assert.AreEqual(eDrawingColorType.Rgb, pg1[0].Bullet.Color.ColorType);
                Assert.IsNotNull(pg1[0].Bullet.Color.RgbColor);
                Assert.AreEqual("Wingdings", pg1[0].Bullet.Font.Typeface);
                Assert.AreEqual("05000000000000000000", pg1[0].Bullet.Font.Panose);
                Assert.AreEqual(ePitchFamily.Variable, pg1[0].Bullet.Font.PitchFamily);
                Assert.AreEqual(2, pg1[0].Bullet.Font.Charset);
                Assert.IsTrue(pg1[2].DefaultRunProperties.IsEmpty);


                var pg2 = ws1.Drawings[1].As.Shape.TextBody.Paragraphs;
                Assert.AreEqual(2, pg2[0].TabStops.Count);
                Assert.AreEqual(eTabStopParagraphAlignment.Decimal, pg2[0].TabStops[0].Alignment);
            }
        }
        [TestMethod]
        public void AddParagraphsToShapes()
        {
            using (var p = OpenPackage("AddParagraphs.xlsx", true))
            {
                var ws1 = p.Workbook.Worksheets.Add("sheet1");
                var shp = ws1.Drawings.AddShape("Shape1", eShapeStyle.Rect);
                shp.TextBody.TextAutofit = eTextAutofit.ShapeAutofit;
                shp.TextBody.Anchor = eTextAnchoringType.Top;
                Assert.AreEqual(0, shp.RichText.Count);
                var pg1 = shp.TextBody.Paragraphs.Add("Paragraph");
                pg1.HorizontalAlignment = eTextAlignment.Center;
                var tr1 = pg1.TextRuns[0];
                var tr2 = pg1.TextRuns.Add(" 1");
                tr1.Fill.Color = Color.Green;
                tr2.Fill.Color = Color.Red;

                var pg2 = shp.TextBody.Paragraphs.Add("This is paragraph 2");
                pg2.HorizontalAlignment = eTextAlignment.Right;
                pg2.TextRuns[0].FontSize = 18;
                pg2.TextRuns[0].HighlightColor.SetPresetColor(ePresetColor.Aqua);
                pg1.DefaultRunProperties.LatinFont = "Arial";
                Assert.AreEqual("Paragraph 1\r\nThis is paragraph 2", shp.Text);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ReadParagraphsCharts()
        {
            using (var p = OpenTemplatePackage("ChartForSvg.xlsx"))
            {
                var ws1 = p.Workbook.Worksheets[0];
                var pg = ws1.Drawings[1].As.Chart.LineChart.Title.TextBody.Paragraphs;

                Assert.IsNull(pg[0].DefaultTabSize);
                Assert.AreEqual(eTextAlignment.Right, pg[0].HorizontalAlignment);
                Assert.AreEqual(18, pg[0].DefaultRunProperties.Size);
                Assert.IsFalse(pg[0].DefaultRunProperties.Italic);
                Assert.AreEqual(OfficeOpenXml.Style.eDrawingTextLineSpacing.Single, pg[0].LineSpacing.LineSpacingType);
                Assert.AreEqual(100D, pg[0].LineSpacing.Value);
                Assert.AreEqual(0D, pg[0].SpaceAfter.Value);
                Assert.AreEqual(0D, pg[0].SpaceBefore.Value);
                Assert.IsFalse(pg[0].DefaultRunProperties.IsEmpty);
                Assert.IsTrue(pg[1].DefaultRunProperties.IsEmpty);
            }
        }

        /// <summary>
        /// Design principle: The default font stays.
        /// If you set a default paragraph Font then un-specified fonts will fall-back to that font
        /// The last used font of a previous text-run is NOT automatically applied as Excel might.
        /// This as if unspecified for a text-run it is assumed a user would want to use the default font for the paragraph.
        /// </summary>
        [TestMethod]
        public void TextInShape()
        {
            using (var p = OpenPackage("TestTextInShapeNew.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("ShapeWorksheet");

                var sunShape = ws.Drawings.AddShape("Sun", eShapeStyle.Sun);

                sunShape.Font.SetFromFont("Calibri", 14, true);

                sunShape.SetSize(500, 500);

                var rt1 = sunShape.RichText.Add("Text One", true);

                var rt2 = sunShape.RichText.Add("SubText One", false);

                rt2.LatinFont = "Algerian";

                var latinFontForRun = rt2.LatinFont;

                var rt3 = sunShape.RichText.Add("SubText two", false);

                var rt21 = sunShape.RichText.Add("Text Two", true);

                Assert.AreEqual("Calibri", rt21.LatinFont);
                Assert.AreEqual(14f, rt21.Size);

                Assert.AreEqual("Calibri", sunShape.TextBody.Paragraphs[1].DefaultRunProperties.LatinFont);
                Assert.AreEqual(14, sunShape.TextBody.Paragraphs[1].DefaultRunProperties.Size);

                var rt22 = sunShape.RichText.Add("\r\n Subtext TwoOne \r\n", false);

                var size = rt22.Size;

                rt22.LatinFont = "Algerian";
                rt22.Size = 12;
                rt22.Color = Color.Red;
                rt22.Bold = false;
                rt22.Italic = true;

                Assert.AreEqual("Algerian", rt22.LatinFont);
                Assert.AreEqual(12f, rt22.Size);

                var rt23 = sunShape.RichText.Add("Subtext TwoTwo", false);

                Assert.AreEqual("Calibri", rt23.LatinFont);
                Assert.AreEqual(14f, rt23.Size);

                SaveAndCleanup(p);
            }
        }
    }
}
