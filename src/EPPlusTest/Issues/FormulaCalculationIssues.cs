using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace EPPlusTest.Issues
{
	[TestClass]
	public class FormulaCalculationIssues : TestBase
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
		public void I1228()
		{
			using (var pck = new ExcelPackage())
			{
				using (var pckTemplate = OpenTemplatePackage("MyIssue.xlsx"))
				{
					pck.Workbook.Worksheets.Add("Foo", pckTemplate.Workbook.Worksheets[1]);
				}

				pck.Workbook.Calculate(x => x.AllowCircularReferences = true);
			}
		}
		[TestMethod]
		public void I1229()
		{
			using (var p = OpenPackage("XLOOKUP.xlsx"))
			{
				var ws = p.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1:A5"].Formula = "XLOOKUP(B1,$C$1:$C$5,$D$1:$D$5,0)";
				ws.Cells["E1"].Formula = "XLOOKUP(B1:B5,$C$1:$C$5,$D$1:$D$5,0)";

				ws.Cells["B1"].Value = 1;
				ws.Cells["B2"].Value = 2;
				ws.Cells["B3"].Value = 3;
				ws.Cells["B4"].Value = 5;
				ws.Cells["B5"].Value = 4;

				ws.Cells["C1"].Value = 1;
				ws.Cells["C2"].Value = 2;
				ws.Cells["C3"].Value = 3;
				ws.Cells["C4"].Value = 5;
				ws.Cells["C5"].Value = 4;

				ws.Cells["D1"].Value = 10;
				ws.Cells["D2"].Value = 12;
				ws.Cells["D3"].Value = 13;
				ws.Cells["D4"].Value = 14;
				ws.Cells["D5"].Value = 15;


				p.Workbook.Calculate();

				Assert.AreEqual(10, ws.Cells["A1"].Value);
				Assert.AreEqual(12, ws.Cells["A2"].Value);
				Assert.AreEqual(13, ws.Cells["A3"].Value);
				Assert.AreEqual(14, ws.Cells["A4"].Value);
				Assert.AreEqual(15, ws.Cells["A5"].Value);

				Assert.AreEqual(10, ws.Cells["E1"].Value);
				Assert.AreEqual(12, ws.Cells["E2"].Value);
				Assert.AreEqual(13, ws.Cells["E3"].Value);
				Assert.AreEqual(14, ws.Cells["E4"].Value);
				Assert.AreEqual(15, ws.Cells["E5"].Value);

			}
		}
	}
}
