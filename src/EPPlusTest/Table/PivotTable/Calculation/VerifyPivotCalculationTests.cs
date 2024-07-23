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
        private object GetPtData(ExcelPivotTable pt, int datafield, params object[] values)
		{
			var l = new List<PivotDataFieldItemSelection>();
			int ix = 0;
			foreach (var f in pt.RowColumnFieldIndicies)
			{
				if (values!=null && values[ix] != null)
				{
					l.Add(new PivotDataFieldItemSelection(pt.Fields[f].Name, values[ix]));
				}
				ix++;
			}

			return pt.GetPivotData(pt.DataFields[datafield].Name, l);
		}
	}
}
