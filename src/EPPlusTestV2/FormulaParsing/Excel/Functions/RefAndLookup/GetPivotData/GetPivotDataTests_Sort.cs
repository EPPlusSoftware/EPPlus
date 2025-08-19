using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using OfficeOpenXml.Table.PivotTable.Calculation;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
	[TestClass]
	public class GetPivotDataTests_Sort : TestBase
	{
		private static ExcelWorksheet _dateWs1, _dateWs2, _dateWs3;
		private static ExcelPackage _package;
		[ClassInitialize]
		public static void TestInitialize(TestContext context)
		{
			_package = OpenPackage("GetPivotData_Sort.xlsx", true);
			_dateWs1 = _package.Workbook.Worksheets.Add("Data1");
			_dateWs2 = _package.Workbook.Worksheets.Add("Data2");
			_dateWs3 = _package.Workbook.Worksheets.Add("Data3");
			LoadItemData(_dateWs1);
			LoadTestdata(_dateWs2);
			LoadDateAndTime(_dateWs3);
		}

		private static void LoadDateAndTime(ExcelWorksheet ws)
		{
			ws.Cells["K1"].Value = "Item";
			ws.Cells["L1"].Value = "Category";
			ws.Cells["M1"].Value = "Stock";
			ws.Cells["N1"].Value = "Price";
			ws.Cells["O1"].Value = "Date for grouping";

			ws.Cells["K2"].Value = "Crowbar";
			ws.Cells["L2"].Value = "Hardware";
			ws.Cells["M2"].Value = 12;
			ws.Cells["N2"].Value = 85.2;
			ws.Cells["O2"].Value = new DateTime(2010, 1, 31, 12, 30, 3);

			ws.Cells["K3"].Value = "Crowbar";
			ws.Cells["L3"].Value = "Hardware";
			ws.Cells["M3"].Value = 15;
			ws.Cells["N3"].Value = 12.2;
			ws.Cells["O3"].Value = new DateTime(2010, 2, 28, 5, 40, 59);

			ws.Cells["K4"].Value = "Hammer";
			ws.Cells["L4"].Value = "Hardware";
			ws.Cells["M4"].Value = 550;
			ws.Cells["N4"].Value = 72.7;
			ws.Cells["O4"].Value = new DateTime(2010, 3, 31, 16, 1, 20);

			ws.Cells["K5"].Value = "Hammer";
			ws.Cells["L5"].Value = "Hardware";
			ws.Cells["M5"].Value = 120;
			ws.Cells["N5"].Value = 11.3;
			ws.Cells["O5"].Value = new DateTime(2010, 4, 30, 18, 32, 55);

			ws.Cells["K6"].Value = "Crowbar";
			ws.Cells["L6"].Value = "Hardware";
			ws.Cells["M6"].Value = 120;
			ws.Cells["N6"].Value = 173.2;
			ws.Cells["O6"].Value = new DateTime(2010, 5, 31, 0, 14, 33);

			ws.Cells["K7"].Value = "Hammer";
			ws.Cells["L7"].Value = "Hardware";
			ws.Cells["M7"].Value = 1;
			ws.Cells["N7"].Value = 4.2;
			ws.Cells["O7"].Value = new DateTime(2010, 6, 30, 19, 57, 01);

			ws.Cells["K8"].Value = "Saw";
			ws.Cells["L8"].Value = "Hardware";
			ws.Cells["M8"].Value = 4;
			ws.Cells["N8"].Value = 33.12;
			ws.Cells["O8"].Value = new DateTime(2010, 6, 28, 05, 28, 29);

			ws.Cells["K9"].Value = "Screwdriver";
			ws.Cells["L9"].Value = "Hardware";
			ws.Cells["M9"].Value = 1200;
			ws.Cells["N9"].Value = 45.2;
			ws.Cells["O9"].Value = new DateTime(2010, 8, 31, 22, 47, 23);

			ws.Cells["K10"].Value = "Apple";
			ws.Cells["L10"].Value = "Groceries";
			ws.Cells["M10"].Value = 807;
			ws.Cells["N10"].Value = 1.2;
			ws.Cells["O10"].Value = new DateTime(2010, 9, 30, 2, 5, 46);

			ws.Cells["K11"].Value = "Butter";
			ws.Cells["L11"].Value = "Groceries";
			ws.Cells["M11"].Value = 52;
			ws.Cells["N11"].Value = 7.2;
			ws.Cells["O11"].Value = new DateTime(2010, 10, 31, 10, 12, 52);

			ws.Cells["K12"].Value = "Monkey Wrench";
			ws.Cells["L12"].Value = "Hardware";
			ws.Cells["M12"].Value = 5;
			ws.Cells["N12"].Value = 233;
			ws.Cells["O12"].Value = new DateTime(2011, 1, 31, 16, 33, 8);

			ws.Cells["O2:O12"].Style.Numberformat.Format = "yyyy-MM-dd";
		}

		[ClassCleanup]
		public static void TestCleanup()
		{
			SaveAndCleanup(_package);
			_package.Dispose();
		}
		[TestMethod]
		public void GetPivotData_Sort_Date_Descending()
		{
			var ws = _package.Workbook.Worksheets.Add("SortDescending");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs1.Cells["K1:O11"], "PivotTable1");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.Sort = eSortType.Descending;
			var df = pt.DataFields.Add(pt.Fields["Stock"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);			
			ws.Calculate();

			//Assert.AreEqual(2881D, (double)ws.Cells["G5"].Value);
			//Assert.AreEqual(2881D, (double)ws.Cells["G6"].Value);
		}
		[TestMethod]
		public void GetPivotData_Sort_Date_Ascending()
		{
			var ws = _package.Workbook.Worksheets.Add("SortAscending");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs1.Cells["K1:O11"], "PivotTable1");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.Sort = eSortType.Ascending;
			var df = pt.DataFields.Add(pt.Fields["Stock"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			ws.Calculate();

			//Assert.AreEqual(2881D, (double)ws.Cells["G5"].Value);
			//Assert.AreEqual(2881D, (double)ws.Cells["G6"].Value);
		}

	}
}
