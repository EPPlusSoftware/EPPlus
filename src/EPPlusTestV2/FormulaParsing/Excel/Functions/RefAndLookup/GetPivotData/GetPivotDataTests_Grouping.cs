using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
	[TestClass]
	public class GetPivotDataTests_Grouping : TestBase
	{
		private static ExcelWorksheet _dateWs1, _dateWs2, _dateWs3;
		private static ExcelPackage _package;
		[ClassInitialize]
		public static void TestInitialize(TestContext context)
		{
			_package = OpenPackage("GetPivotData_Grouping.xlsx", true);
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
		public void GetPivotData_Grouping_Year()
		{
			var ws = _package.Workbook.Worksheets.Add("DateGroup_Year");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs1.Cells["K1:O11"], "PivotTable1");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.AddDateGrouping(eDateGroupBy.Years);
			var df = pt.DataFields.Add(pt.Fields["Stock"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Date for grouping\",2010)";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Stock\",$A$1)";
			ws.Calculate();

			Assert.AreEqual(2881D, (double)ws.Cells["G5"].Value);
			Assert.AreEqual(2881D, (double)ws.Cells["G6"].Value);
		}
		[TestMethod]
		public void GetPivotData_Grouping_YearMonth()
		{
			var ws = _package.Workbook.Worksheets.Add("DateGroup_YearMonth");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs1.Cells["K1:O11"], "PivotTable2");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months);
			var df = pt.DataFields.Add(pt.Fields["Stock"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Date for grouping\", 3)";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Stock\",$A$1)";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Stock\",$A$1, \"Date for grouping\", 5, \"Years\", 2010)";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Date for grouping\",\"AUG\",\"Years\",2010)";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Date for grouping\",\"sep\")";
			ws.Calculate();

			Assert.AreEqual(550D, (double)ws.Cells["G5"].Value);
			Assert.AreEqual(2881D, (double)ws.Cells["G6"].Value);
			Assert.AreEqual(120D, (double)ws.Cells["G7"].Value);
			Assert.AreEqual(1200D, (double)ws.Cells["G8"].Value);
			Assert.AreEqual(807D, (double)ws.Cells["G9"].Value);
		}
		[TestMethod]
		public void GetPivotData_Grouping_YearMonthDay()
		{
			SwitchToCulture();
			var ws = _package.Workbook.Worksheets.Add("DateGroup_YearMonthDay");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs1.Cells["K1:O11"], "PivotTable3");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days);
			var df = pt.DataFields.Add(pt.Fields["Stock"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Date for grouping\", 31)";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Date for grouping\", 31+28)";
			ws.Cells["G7"].Formula = "=GETPIVOTDATA(\"Stock\",$A$1,\"Date for grouping\",121,\"Months\",4,\"Years\",2010)";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Stock\",$A$1, \"Date for grouping\", \"28-Jun\", \"Months\", 6, \"Years\", 2010)";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Stock\",$A$1, \"Date for grouping\", \"30-Jun\", \"Years\", 2010)";
			ws.Cells["G10"].Formula = "=GETPIVOTDATA(\"Stock\",$A$1,\"Years\",2010)";
			ws.Cells["G11"].Formula = "=GETPIVOTDATA(\"Stock\",$A$1)";
			ws.Calculate();

			Assert.AreEqual(12D, (double)ws.Cells["G5"].Value);
			Assert.AreEqual(15D, (double)ws.Cells["G6"].Value);
			Assert.AreEqual(120D, (double)ws.Cells["G7"].Value);
			Assert.AreEqual(4D, (double)ws.Cells["G8"].Value);
			Assert.AreEqual(1D, (double)ws.Cells["G9"].Value);
			Assert.AreEqual(ErrorValues.RefError, ws.Cells["G10"].Value);
			Assert.AreEqual(2881D, (double)ws.Cells["G11"].Value);
			SwitchBackToCurrentCulture();
		}
		[TestMethod]
		public void GetPivotData_Grouping_Hour()
		{
			var ws = _package.Workbook.Worksheets.Add("DateGroup_Hours");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs3.Cells["K1:O12"], "PivotTable4");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.Name = "Hours";
			rf.AddDateGrouping(eDateGroupBy.Hours);
			var df = pt.DataFields.Add(pt.Fields["Stock"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 12)";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 16)";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 22)";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Stock\",$A$1)";

			ws.Calculate();

			Assert.AreEqual(12D, (double)ws.Cells["G5"].Value);
			Assert.AreEqual(555D, (double)ws.Cells["G6"].Value);
			Assert.AreEqual(1200D, (double)ws.Cells["G7"].Value);
			Assert.AreEqual(2886D, (double)ws.Cells["G8"].Value);
		}
		[TestMethod]
		public void GetPivotData_Grouping_HourMinuts()
		{
			var ws = _package.Workbook.Worksheets.Add("DateGroup_HoursMinutes");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs3.Cells["K1:O12"], "PivotTable5");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.Name = "Minutes";
			rf.AddDateGrouping(eDateGroupBy.Hours | eDateGroupBy.Minutes);
			var df = pt.DataFields.Add(pt.Fields["Stock"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			var grandTotal = pt.CalculatedData.GetValue("Stock");

			var item1230 = pt.CalculatedData
                .SelectField("Hours", 12)
                .SelectField("Minutes", 30)
                .GetValue("Stock");

            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 12, \"Minutes\", 30)";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 16)";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 22)";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Stock\",$A$1)";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 16, \"Minutes\", 1)";

			ws.Calculate();

			Assert.AreEqual(12D, (double)ws.Cells["G5"].Value);
			Assert.AreEqual(ErrorValues.RefError, ws.Cells["G6"].Value);
			Assert.AreEqual(1200D, (double)ws.Cells["G7"].Value);
			Assert.AreEqual(2886D, (double)ws.Cells["G8"].Value);
			Assert.AreEqual(550D, (double)ws.Cells["G9"].Value);
		}
		[TestMethod]
		public void GetPivotData_Grouping_HourMinutsSeconds()
		{
			var ws = _package.Workbook.Worksheets.Add("DateGroup_HoursMinutesSeconds");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs3.Cells["K1:O12"], "PivotTable6");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.Name = "Seconds";
			rf.AddDateGrouping(eDateGroupBy.Hours | eDateGroupBy.Minutes | eDateGroupBy.Seconds);
			var df = pt.DataFields.Add(pt.Fields["Stock"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			//ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 12, \"Minutes\", 30)";
			//ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 16, \"Seconds\",20)";
			//ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Seconds\", 33)";
			//ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Stock\",$A$1)";
			//ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 16, \"Minutes\", 1, \"Seconds\", 20)";
			ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Stock\",$A$1,\"Hours\", 16, \"Minutes\", 1, \"Seconds\", 59)";

			ws.Calculate();

			//Assert.AreEqual(12D, (double)ws.Cells["G5"].Value);
			//Assert.AreEqual(550D, (double)ws.Cells["G6"].Value);
			//Assert.AreEqual(120D, (double)ws.Cells["G7"].Value);
			//Assert.AreEqual(2886D, (double)ws.Cells["G8"].Value);
			//Assert.AreEqual(550D, (double)ws.Cells["G9"].Value);
			Assert.AreEqual(ErrorValues.RefError, ws.Cells["G10"].Value);
		}
		[TestMethod]
		public void GetPivotData_Grouping_Numbers()
		{
			var ws = _package.Workbook.Worksheets.Add("NumberGroup");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs2.Cells["A1:D100"], "PivotTable7");
			var rf = pt.RowFields.Add(pt.Fields["NumValue"]);
			rf.AddNumericGrouping(0, 100, 10);
			var df = pt.DataFields.Add(pt.Fields["NumFormattedValue"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",0)";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",20)";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",60)";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",90)";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1)";
			ws.Cells["G10"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",200)";
			ws.Cells["G11"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",17)";

			ws.Calculate();

			Assert.AreEqual(1452D, (double)ws.Cells["G5"].Value);
			Assert.AreEqual(8085D, (double)ws.Cells["G6"].Value);
			Assert.AreEqual(21285D, (double)ws.Cells["G7"].Value);
			Assert.AreEqual(34485D, (double)ws.Cells["G8"].Value);
			Assert.AreEqual(166617D, (double)ws.Cells["G9"].Value);
			Assert.AreEqual(ErrorValues.RefError, ws.Cells["G10"].Value);
			Assert.AreEqual(ErrorValues.RefError, ws.Cells["G11"].Value);
		}
		[TestMethod]
		public void GetPivotData_Grouping_Numbers_Decimals()
		{
			var ws = _package.Workbook.Worksheets.Add("NumberGroupDecimals");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs2.Cells["A1:D100"], "PivotTable7");
			var rf = pt.RowFields.Add(pt.Fields["NumValue"]);
			rf.AddNumericGrouping(0, 100, 15.55);
			var df = pt.DataFields.Add(pt.Fields["NumFormattedValue"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",0)";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",15.55)";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",77.75)";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",93.3)";

			ws.Calculate();

			Assert.AreEqual(3927D, (double)ws.Cells["G5"].Value);
			Assert.AreEqual(12408D, (double)ws.Cells["G6"].Value);
			Assert.AreEqual(45144D, (double)ws.Cells["G7"].Value);
			Assert.AreEqual(22407D, (double)ws.Cells["G8"].Value);
		}
		[TestMethod]
		public void GetPivotData_Grouping_Numbers_Intervall_Over_And_Under()
		{
			var ws = _package.Workbook.Worksheets.Add("NumberGroup_OverUnder");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs2.Cells["A1:D100"], "PivotTable7");
			var rf = pt.RowFields.Add(pt.Fields["NumValue"]);
			rf.AddNumericGrouping(30, 70, 10);
			var df = pt.DataFields.Add(pt.Fields["NumFormattedValue"]);
			df.Function = DataFieldFunctions.Sum;
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",\"<\")";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",30)";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",70)";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1,\"NumValue\",\">\")";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"NumFormattedValue\",$A$1)";

			ws.Calculate();

			Assert.AreEqual(14322D, (double)ws.Cells["G5"].Value);
			Assert.AreEqual(11385D, (double)ws.Cells["G6"].Value);
			Assert.AreEqual(84645D, (double)ws.Cells["G7"].Value);
			Assert.AreEqual(84645D, (double)ws.Cells["G8"].Value);
			Assert.AreEqual(166617D, (double)ws.Cells["G9"].Value);
		}
	}
}