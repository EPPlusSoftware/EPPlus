using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.IO;

namespace EPPlusTest.Drawing
{
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
                var r2=shape.RichText.Add("L", true);
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


                var r3=shape.RichText.Add("Line3", true);
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
                AssertIfNotExists("DrawingRichTextRead.xlsx");
                using (var p = OpenPackage("DrawingRichTextRead.xlsx"))
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
                    var pg1 = shp.TextBody.Paragraphs.Add("Paragraph 1");
                    pg1.DefaultRunProperties.LatinFont = "Aptos Narrow"; 
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

        }
    }
}
