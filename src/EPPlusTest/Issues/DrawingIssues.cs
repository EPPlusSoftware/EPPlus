using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.IO;
using System.Drawing;
namespace EPPlusTest.Issues
{
	[TestClass]
	public class DrawingIssues : TestBase
	{
		[ClassInitialize]
		public static void Init(TestContext context)
		{
		}
		[ClassCleanup]
		public static void Cleanup()
		{
		}
		[TestInitialize]
		public void Initialize()
		{
		}
		[TestMethod]
		public void s633()
		{
			using (var p = OpenTemplatePackage("s633.xlsx"))
			{
				var sheet = p.Workbook.Worksheets[0];
				var pic=sheet.Drawings[0].As.Picture;
			}
		}
        [TestMethod]
        public void i1446()
        {
            using (var p = OpenTemplatePackage("Scrollbar.xlsx"))
            {
                var sheet = p.Workbook.Worksheets[0];
                var sb = sheet.Drawings[0].As.Control.SpinButton;
                var sbr = sheet.Drawings[1].As.Control.ScrollBar;
            }
        }

        [TestMethod]
        public void i1640()
        {
            using (var package = OpenTemplatePackage("i1640.xlsx"))
            {
                package.Workbook.MaxFontWidth = 8;
                FontSize.FontWidths.Add("ＭＳ Ｐゴシック", new Dictionary<float, short> { { 11, 8 } });

                var sheet = package.Workbook.Worksheets[0];
                CopyRows(sheet, 1, 10, 11, 20);
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s762()
        {
            using (var package = OpenTemplatePackage("s762.xlsx"))
            {
                var sheet = package.Workbook.Worksheets[0];
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void i1673()
        {
            using (var package = OpenTemplatePackage("i1673.xlsx"))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets[0];
                ws.Drawings.Count();

                SaveAndCleanup(package);
            }
        }

        public void CopyRows(ExcelWorksheet excelWorksheet, int sourceFrom, int sourceTo, int destFrom, int destTo)
        {
            for (int i = destFrom; i <= destTo; i++)
            {
                excelWorksheet.Row(i).Height = excelWorksheet.Row(sourceFrom + i - destFrom).Height;
            }

            excelWorksheet.Cells[sourceFrom.ToString() + ":" + sourceTo].Copy(
                excelWorksheet.Cells[destFrom.ToString() + ":" + destTo]);
        }
        [TestMethod]
        public void OleTest1()
        {
            using var p = OpenTemplatePackage("OleMSChart.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var oleObject = ws.Drawings[0].As.OleObject;
            var grph = oleObject.GetEmbeddedObjectBytes();
            Assert.IsNotNull(oleObject);
            Assert.IsNotNull(oleObject.ProgId);
            Assert.IsNotNull(oleObject.Image);
            Assert.IsNull(oleObject.ExternalLink);
            SaveAndCleanup(p);
        }
        public void i1902()
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");

            var rect = worksheet.Drawings.AddShape("rect1", eShapeStyle.Rect);

            Console.WriteLine(rect.Text);
        }
        [TestMethod]
        public void CommentIssue()
        {
            using var package = OpenPackage("CommentIssuePosition.xlsx", true);
            var ws = package.Workbook.Worksheets.Add("Sheet1");

            //Add a comment using the Comment collection
            var comment = ws.Comments.Add(ws.Cells["B3"], "This column contains the size of the files.", "JK");
            //This sets the size and position. (The position is only when the comment is visible)
            comment.From.Column = 7;
            comment.From.Row = 3;
            comment.To.Column = 16;
            comment.To.Row = 8;
            comment.BackgroundColor = Color.White;
            comment.RichText.Add("\r\nTo format the numbers use the Numberformat-property like:\r\n");

            Assert.AreEqual("7, 15, 3, 2, 16, 31, 8, 1", comment.Anchor);

            ws.Cells["B3:B42"].Style.Numberformat.Format = "#,##0";

            //Format the code using the RichText Collection
            var rc = comment.RichText.Add("//Format the Size and Count column\r\n");
            rc.FontName = "Courier New";
            rc.Color = Color.FromArgb(0, 128, 0);
            rc = comment.RichText.Add("ws.Cells[");
            rc.Color = Color.Black;
            rc = comment.RichText.Add("\"B3:B42\"");
            rc.Color = Color.FromArgb(123, 21, 21);
            rc = comment.RichText.Add("].Style.Numberformat.Format = ");
            rc.Color = Color.Black;
            rc = comment.RichText.Add("\"#,##0\"");
            rc.Color = Color.FromArgb(123, 21, 21);
            rc = comment.RichText.Add(";");
            rc.Color = Color.Black;

            SaveAndCleanup(package);
        }
        [TestMethod]
        public void s838()
        {
            using var p = OpenTemplatePackage("s838.xlsx");
            Assert.AreEqual(2, p.Workbook.Worksheets[0].Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void ThreeDModel()
        {
            using var p = OpenTemplatePackage("3dmodel.xlsx");
            Assert.AreEqual(1, p.Workbook.Worksheets[0].Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        //i2201 See #2201
        public void EnsureWhiteSpaceIsPreservedInShapes()
        {
            var ms = new MemoryStream();
            using (var origP = new ExcelPackage("whiteSpace.xlsx"))
            {
                var ws = origP.Workbook.Worksheets.Add("newWs");
                var txtBox = ws.Drawings.AddTextbox("txtbox1", " ");
                origP.SaveAs(ms);
            }

            string retText = "";

            using (var readP = new ExcelPackage(ms))
            {
                var myShape = readP.Workbook.Worksheets[0].Drawings[0].As.Shape;
                retText = myShape.Text;
            }

            Assert.AreEqual(" ", retText);
        }
        [TestMethod]
        public void i2278()
        {
            using (var package = OpenTemplatePackage("i2278.xlsx"))
            {
                using var target = new ExcelPackage();
                var targetSheet = target.Workbook.Worksheets.Add("Sheet1");

                var worksheet = package.Workbook.Worksheets[0];

                foreach (var drawing in worksheet.Drawings)
                {
                    drawing.Copy(targetSheet, drawing.From.Row, drawing.From.Column, drawing.From.RowOff, drawing.From.ColumnOff);
                }

                SaveAndCleanup(package); 
            }
        }
        [TestMethod]
        public void i2303()
        {
            using (var package = OpenTemplatePackage("i2303.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.First();
                var drawing = sheet.Drawings.First();
                var image = drawing.As.Picture.Image;
                
                SaveAndCleanup(package);
            }
        }        
    }
}

