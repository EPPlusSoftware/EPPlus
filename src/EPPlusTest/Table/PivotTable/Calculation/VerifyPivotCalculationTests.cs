using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
namespace EPPlusTest.Table.PivotTable.Calculation
{
	[TestClass]
	public class VerifyPivotCalculationTests : TestBase
	{
		static ExcelPackage _package;
		static ExcelWorksheet _ptWs;
		static ExcelWorksheet _ptWs2;
		[ClassInitialize]
		public static void Init(TestContext context)


		{
			InitBase();
			_package = OpenTemplatePackage("GetPivotData\\PivotTableCalcTest.xlsx");
			_ptWs = _package.Workbook.Worksheets["PivotTables"];
		}
		[ClassCleanup]
		public static void Cleanup()
		{
			_package.Dispose();
		}
        [TestMethod]
        public void VerifyCalculationMD()
        {
            var ws = _package.Workbook.Worksheets[3];
            var pt = ws.PivotTables[0];
            pt.Calculate();

            Assert.AreEqual(4335.69, GetPtData(pt, 0, "Good Samaritan Hospital", "Tissue Sealer", "2023", "mar"));
            Assert.AreEqual(34454.62, GetPtData(pt, 0, "Palm Beach Garden Comm Hospital", null, 2022, "Nov"));
        }
        private object GetPtData(ExcelPivotTable pt, int datafield, params object[] values)
		{
			var l = new List<PivotDataCriteria>();
			int ix = 0;
			foreach (var f in pt.RowColumnFieldIndicies)
			{
				if (values!=null && values[ix] != null)
				{
					l.Add(new PivotDataCriteria(pt.Fields[f], values[ix]));
				}
				ix++;
			}

			return pt.GetPivotData(l, pt.DataFields[datafield]);
		}
	}
}
