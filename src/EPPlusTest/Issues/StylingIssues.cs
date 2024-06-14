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
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
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
    }
}
