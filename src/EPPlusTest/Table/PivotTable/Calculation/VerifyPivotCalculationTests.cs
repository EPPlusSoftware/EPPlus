using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using System.Collections.Generic;
using OfficeOpenXml;
<<<<<<< HEAD
=======
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
>>>>>>> develop7
namespace EPPlusTest.Table.PivotTable.Calculation
{
    [TestClass]
	public class VerifyPivotCalculationTests : TestBase
	{
<<<<<<< HEAD
		static ExcelPackage _package;
		static ExcelWorksheet _ptWs;
		static ExcelWorksheet _ptWs2;
		[ClassInitialize]
		public static void Init(TestContext context)


		{
			InitBase();
			_package = OpenTemplatePackage("GetPivotData\\PivotTableCalcTest.xlsx");
			_ptWs = _package.Workbook.Worksheets["PivotTables"];
=======
		[ClassInitialize]
		public static void Init(TestContext context)
		{
			InitBase();
>>>>>>> develop7
		}
		[ClassCleanup]
		public static void Cleanup()
		{
<<<<<<< HEAD
			_package.Dispose();
=======
>>>>>>> develop7
		}
        [TestMethod]
        public void VerifyCalculationMD()
        {
<<<<<<< HEAD
            var ws = _package.Workbook.Worksheets[3];
            var pt = ws.PivotTables[0];

            Assert.AreEqual(4335.69, GetPtData(pt, 0, "Good Samaritan Hospital", "Tissue Sealer", "2023", "mar"));
            Assert.AreEqual(34454.62, GetPtData(pt, 0, "Palm Beach Garden Comm Hospital", null, 2022, "Nov"));
=======
			using (var p = OpenTemplatePackage("GetPivotData\\PivotTableCalcTest.xlsx"))
			{
				var ptWs = p.Workbook.Worksheets["PivotTables"];
				var ws = p.Workbook.Worksheets[3];
				var pt = ws.PivotTables[0];
			}
>>>>>>> develop7
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
