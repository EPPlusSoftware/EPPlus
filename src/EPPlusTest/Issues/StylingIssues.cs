using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
namespace EPPlusTest
{
	[TestClass]
	public class StylingIssues : TestBase
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
		public void i1291()
		{
			using (var p = OpenPackage("i1291.xlsx", true))
			{
				var ws = p.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1"].Style.Font.Name = "+Headings";
				SaveAndCleanup(p);
			}
		}
		[TestMethod]
		public void i1320()
		{
			using(var package = OpenPackage("i1320.xlsx", true))
			{
				var worksheet = package.Workbook.Worksheets.Add("Worksheet");

				// Default entire worksheet to Arial 12pt
				worksheet.Cells["A:XFD"].Style.Font.Name = "Arial";
				worksheet.Cells["A:XFD"].Style.Font.Size = 12;

				// Header row
				worksheet.Row(1).Style.Font.Bold = true;
				worksheet.Cells[1, 1].Value = "COL1";
				worksheet.Cells[1, 2].Value = "COL2";
				worksheet.Cells[1, 3].Value = "COL3";

				Assert.AreEqual("Arial", worksheet.Row(1).Style.Font.Name);
				Assert.AreEqual("Arial", worksheet.Cells[1, 1].Style.Font.Name);
				Assert.AreEqual("Arial", worksheet.Cells[1, 2].Style.Font.Name);
				Assert.AreEqual("Arial", worksheet.Cells[1, 3].Style.Font.Name);

				SaveAndCleanup(package);
			}
		}
		[TestMethod]
		public void i1454()
		{
            using var p1 = OpenTemplatePackage("i1454.xlsx");
			var ws = p1.Workbook.Worksheets[0];
			using var p2 = OpenPackage("i1454-copy.xlsx", true);
            p2.Workbook.Worksheets.Add($"{ws.Name} [2]", ws);
			SaveAndCleanup(p2);
        }
        [TestMethod]
        public void IssueMissingDecimalsTextFormular()
        {
            //Issue: TEXT-formular deletes decimals in german format
            using var p = OpenTemplatePackage("Textformat.xlsx");

            SwitchToCulture("de-DE");
            p.Workbook.Calculate();

            Assert.AreEqual("292.336,30 €", p.Workbook.Worksheets[0].Cells["A1"].Text);
            Assert.AreEqual("292336,300000 €", p.Workbook.Worksheets[0].Cells["A2"].Value);
            Assert.AreEqual("292.336 €", p.Workbook.Worksheets[0].Cells["A3"].Value);

            Assert.AreEqual("292.336,30 €", p.Workbook.Worksheets[0].Cells["A5"].Value);
            Assert.AreEqual("-292336-- €", p.Workbook.Worksheets[0].Cells["A6"].Value);
            //Assert.AreEqual("-292336,--€", p.Workbook.Worksheets[0].Cells["A6"].Value);
            Assert.AreEqual("233.127,25 €)", p.Workbook.Worksheets[0].Cells["A7"].Value);
            Assert.AreEqual("-233127--€)", p.Workbook.Worksheets[0].Cells["A8"].Value);
            //Assert.AreEqual("-233127,--€)", p.Workbook.Worksheets[0].Cells["A8"].Value);
            Assert.AreEqual("0,00 €", p.Workbook.Worksheets[0].Cells["A9"].Value);
            Assert.AreEqual("--- €", p.Workbook.Worksheets[0].Cells["A10"].Value);
            //Assert.AreEqual("-,-- €", p.Workbook.Worksheets[0].Cells["A10"].Value);
            Assert.AreEqual("0,00 €)", p.Workbook.Worksheets[0].Cells["A11"].Value);
            Assert.AreEqual("---€)", p.Workbook.Worksheets[0].Cells["A12"].Value);
            //Assert.AreEqual("-,--€)", p.Workbook.Worksheets[0].Cells["A12"].Value);
            Assert.AreEqual("1.027,60 €", p.Workbook.Worksheets[0].Cells["A13"].Value);
            Assert.AreEqual("-1028-- €)", p.Workbook.Worksheets[0].Cells["A14"].Value);
            //Assert.AreEqual("-1028,--€)", p.Workbook.Worksheets[0].Cells["A14"].Value);
            Assert.AreEqual("445,58 €)", p.Workbook.Worksheets[0].Cells["A15"].Value);
            Assert.AreEqual("-446-- €)", p.Workbook.Worksheets[0].Cells["A16"].Value);
            //Assert.AreEqual("-446,--€)", p.Workbook.Worksheets[0].Cells["A16"].Value);
            Assert.AreEqual("0,00 €", p.Workbook.Worksheets[0].Cells["A17"].Value);
            Assert.AreEqual("0,00 €)", p.Workbook.Worksheets[0].Cells["A18"].Value);
            Assert.AreEqual("--- €)", p.Workbook.Worksheets[0].Cells["A19"].Value);
            //Assert.AreEqual("-,--€)", p.Workbook.Worksheets[0].Cells["A19"].Value);
            Assert.AreEqual("--- €", p.Workbook.Worksheets[0].Cells["A20"].Value);
			//Assert.AreEqual("-,--€", p.Workbook.Worksheets[0].Cells["A20"].Value);

			SwitchBackToCurrentCulture();
        }
        [TestMethod]
        public void Issue1493()
        {
            ExcelPackageSettings.CultureSpecificBuildInNumberFormats.Add("de-DE",
                new Dictionary<int, string>()
                {
                   {14, "dd.mm.yyyy"}, {15,"dd. mmm yy"}, {16,"dd. mmm"}, {17,"mmm yy"}, {18, "hh:mm AM/PM" }, {22, "dd.mm.yyyy hh:mm"},{39, "#,##0.00;-#,##0.00"}, {47, "mm:ss,f"}
                });

            using var p = OpenTemplatePackage("i1493.xlsx");
            
            SwitchToCulture("de-DE");
            p.Workbook.NumberFormatToTextHandler = TextHandler;
            p.Workbook.Calculate();
            var ws = p.Workbook.Worksheets[0];

            Assert.AreEqual("123456789,1", ws.Cells["A2"].Text); // actual "123456789,123456"
            Assert.AreEqual("123456789", ws.Cells["A3"].Text);
            Assert.AreEqual("123456789,12", ws.Cells["A4"].Text);
            Assert.AreEqual("123.456.789", ws.Cells["A5"].Text);
            Assert.AreEqual("123.456.789,12", ws.Cells["A6"].Text);
            Assert.AreEqual("123.456.789,12", ws.Cells["A9"].Text);
            Assert.AreEqual("123.456.789,12", ws.Cells["A10"].Text);
            Assert.AreEqual("123.456.789 €", ws.Cells["A11"].Text);
            Assert.AreEqual("123.456.789 €", ws.Cells["A12"].Text);
            Assert.AreEqual("123.456.789,12 €", ws.Cells["A13"].Text);
            Assert.AreEqual("123.456.789,12 €", ws.Cells["A14"].Text);
            Assert.AreEqual("12345678912%", ws.Cells["A15"].Text);
            Assert.AreEqual("12345678912,35%", ws.Cells["A16"].Text);
            Assert.AreEqual("1,23E+08", ws.Cells["A17"].Text);
            Assert.AreEqual("123,5E+6", ws.Cells["A18"].Text); // actual "123456789,1"
            Assert.AreEqual("123456789 1/8", ws.Cells["A19"].Text);
            Assert.AreEqual("123456789 10/81", ws.Cells["A20"].Text);
            Assert.AreEqual("29.03.2018", ws.Cells["A21"].Text);

#if (NET6_0_OR_GREATER)
            Assert.AreEqual("29. März 18", ws.Cells["A22"].Text); // actual "29-März-18"
            Assert.AreEqual("29. März", ws.Cells["A23"].Text); // actual "29-März"
            Assert.AreEqual("Mär 18", ws.Cells["A24"].Text); // actual "Mär-18"

            Assert.AreEqual("Mär 2019", ws.Cells["A38"].Text); // actual "Mär 2019"
#else
            Assert.AreEqual("29. Mrz 18", ws.Cells["A22"].Text); // actual "29-März-18"
            Assert.AreEqual("29. Mrz", ws.Cells["A23"].Text); // actual "29-März"
            Assert.AreEqual("Mrz 18", ws.Cells["A24"].Text); // actual "Mär-18"

            Assert.AreEqual("Mrz 2019", ws.Cells["A38"].Text); // actual "Mär 2019"
            Assert.AreEqual("Samstag, 30. März 2019", ws.Cells["A39"].Text);
#endif            
            Assert.AreEqual("10:45 AM", ws.Cells["A25"].Text); // actual "10:45"
            Assert.AreEqual("10:45:00 AM", ws.Cells["A26"].Text); // actual "10:45:00" 
            Assert.AreEqual("10:45", ws.Cells["A27"].Text);
            Assert.AreEqual("10:45:00", ws.Cells["A28"].Text);
            Assert.AreEqual("29.03.2019 10:45", ws.Cells["A29"].Text); // actual "3.29.19 10:45"
            Assert.AreEqual("44:59", ws.Cells["A30"].Text); // actual "03:59"
            Assert.AreEqual("44:59,9", ws.Cells["A31"].Text); // actual "0359.0"
            Assert.AreEqual("43555,48958", ws.Cells["A32"].Text); // actual "43555,4895832755"
            Assert.AreEqual("1045332:44:59", ws.Cells["A33"].Text); // actual "12:03:59"
            Assert.AreEqual("123.456.789 ", ws.Cells["A35"].Text); // actual "123.456.789"
            Assert.AreEqual("Samstag, 30. März 2019", ws.Cells["A39"].Text);

            Assert.AreEqual("-123.456.789,12", ws.Cells["B9"].Text); // actual "(123.456.789,12)"
            Assert.AreEqual("-123.456.789 €", ws.Cells["B10"].Text); // actual "-123.456.789 €"
            Assert.AreEqual("-1,23E+08", ws.Cells["B17"].Text);
            Assert.AreEqual("-123,5E+6", ws.Cells["B18"].Text); //actual "-123456789,1"
            Assert.AreEqual("-123456789 1/8", ws.Cells["B19"].Text);   //actual: "--123456789 1/" 
            Assert.AreEqual("-123456789 10/81", ws.Cells["B20"].Text); //actual: "--123456789  1/"  

            Assert.AreEqual("0", ws.Cells["C2"].Text);
            Assert.AreEqual("0", ws.Cells["C3"].Text);
            Assert.AreEqual("0,00", ws.Cells["C4"].Text);
            Assert.AreEqual("0,00", ws.Cells["C9"].Text);
            Assert.AreEqual("0,00 €", ws.Cells["C11"].Text);
            Assert.AreEqual("0,00 €", ws.Cells["C13"].Text);
            Assert.AreEqual("0,00 €", ws.Cells["C14"].Text);
            Assert.AreEqual("0%", ws.Cells["C15"].Text);
            Assert.AreEqual("0,00%", ws.Cells["C16"].Text);
            Assert.AreEqual("0,00E+00", ws.Cells["C17"].Text);
            Assert.AreEqual("000,0E+0", ws.Cells["C18"].Text); // actual "0,0"

            Assert.AreEqual("- €", ws.Cells["C34"].Text);
            Assert.AreEqual("- ", ws.Cells["C35"].Text);
            Assert.AreEqual("- €", ws.Cells["C36"].Text);
            Assert.AreEqual("- ", ws.Cells["C37"].Text);
            
            SwitchBackToCurrentCulture();
        }
        [TestMethod]
        public void i1523()
        {
            using (var package = new ExcelPackage())
            {
                SwitchToCulture();
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = 123;
                sheet.Cells["A1"].Style.Numberformat.Format = "General\\ \"mm\"";
                Assert.AreEqual("123 mm", sheet.Cells[1, 1].Text);

                sheet.Cells["A2"].Value = 123456789.1234;
                sheet.Cells["A2"].Style.Numberformat.Format = "General\\ \"mm\"";
                Assert.AreEqual("123456789.1 mm", sheet.Cells[2, 1].Text);
                SwitchBackToCurrentCulture();
            }
        }
        [TestMethod]
        public void s763_1()
        {
            using (ExcelPackage p = OpenPackage("s763.xlsx", true))
            {
                var wb = p.Workbook;
                var decimalList = new List<decimal>();
                decimalList = Enumerable.Range(1, 10).Select(i => (decimal)new Random().NextDouble() * 100000).ToList();
                var ws = wb.Worksheets.Add("NumberFormat");
                var row = 1;
                foreach (var n in decimalList)
                {
                    ws.Cells[row++, 1].Value = n;
                }
                ws.Column(1).Style.Numberformat.Format = "#,#0.00";

                Assert.AreEqual(1, ws.GetStyleInner(1, 1));
                Assert.AreEqual(1, ws.GetStyleInner(10, 1));
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s763_2()
        {
            using (ExcelPackage p = OpenPackage("s763_2.xlsx", true))
            {
                var wb = p.Workbook;
                var decimalList = new List<decimal>();
                decimalList = Enumerable.Range(1, 10).Select(i => (decimal)new Random().NextDouble() * 100000).ToList();
                var ws = wb.Worksheets.Add("NumberFormat");
                var col = 1;
                foreach (var n in decimalList)
                {
                    ws.Cells[1, col++].Value = n;
                }
                ws.Row(1).Style.Numberformat.Format = "#,#0.00";

                Assert.AreEqual(1, ws.GetStyleInner(1, 1));
                Assert.AreEqual(1, ws.GetStyleInner(1, 10));
                SaveAndCleanup(p);
            }
        }
        public string TextHandler(NumberFormatToTextArgs options)
        {
            switch(options.NumberFormat.NumFmtId)
            {
                case 15:
                    break;
            }
            return options.Text;
        }
    }
}