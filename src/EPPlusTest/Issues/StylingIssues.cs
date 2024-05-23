using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
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
	}
}
