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

    }
}

