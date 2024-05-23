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
	public class GetPivotDataTests_CalculatedFields : TestBase
	{
		private static ExcelWorksheet _dateWs1, _dateWs2, _dateWs3;
		private static ExcelPackage _package;
		[ClassInitialize]
		public static void TestInitialize(TestContext context)
		{
			_package = OpenPackage("GetPivotData_CalculatedFields.xlsx", true);
			_dateWs1 = _package.Workbook.Worksheets.Add("Data1");
			LoadItemData(_dateWs1);
		}

		[ClassCleanup]
		public static void TestCleanup()
		{
			SaveAndCleanup(_package);
			_package.Dispose();
		}
		[TestMethod]
		public void GetPivotData_AddCalculatedField()
		{
			var ws = _package.Workbook.Worksheets.Add("SortDescending");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _dateWs1.Cells["K1:O11"], "PivotTable1");			
			var rf = pt.RowFields.Add(pt.Fields[4]);
			pt.Fields.AddCalculatedField("Calculated Value", "Price * Stock * 'Date for grouping'");
			var df = pt.DataFields.Add(pt.Fields["Calculated Value"]);
			df.Function = DataFieldFunctions.Sum;
			ws.Calculate();
		}

	}
}
