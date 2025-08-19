using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
namespace EPPlusTest.Table.PivotTable.Calculation
{
    [TestClass]
	public class VerifyPivotCalculationTests : TestBase
	{
		[ClassInitialize]
		public static void Init(TestContext context)
		{
			InitBase();
		}
		[ClassCleanup]
		public static void Cleanup()
		{
		}
        [TestMethod]
        public void VerifyCalculationMD()
        {
			using (var p = OpenTemplatePackage("GetPivotData\\PivotTableCalcTest.xlsx"))
			{
				var ptWs = p.Workbook.Worksheets["PivotTables"];
				var ws = p.Workbook.Worksheets[3];
				var pt = ws.PivotTables[0];
			}
        }
        [TestMethod]
        public void VerifyDuplicateNumberAndString()
		{
            using (var p = OpenPackage("PivotDupTest.xlsx", true))
			{
				var ws = p.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1"].Value = "Col";
                ws.Cells["B1"].Value = "Value";
                ws.Cells["A2"].Value = "200";
                ws.Cells["B2"].Value = 1;
                ws.Cells["A3"].Value = 200;
                ws.Cells["B3"].Value = 2;
                ws.Cells["A3"].Value = 201;
                ws.Cells["B3"].Value = 3;

				var pt = ws.PivotTables.Add(ws.Cells["E5"], ws.Cells["A1:B3"], "PivotTable1");

				pt.Calculate(true);

				SaveAndCleanup(p);
            }
        }
	}
}
