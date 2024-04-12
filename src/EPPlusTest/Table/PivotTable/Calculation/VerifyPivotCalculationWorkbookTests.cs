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
	public class VerifyPivotCalculationWorkbookTests : TestBase
	{
		static ExcelPackage _package;
		static ExcelWorksheet _ptWs;
		static ExcelWorksheet _ptWs2;
		[ClassInitialize]
		public static void Init(TestContext context)


		{
			InitBase();
			_package = OpenTemplatePackage("PivotTableCalculation.xlsx");
			_ptWs = _package.Workbook.Worksheets["PivotTables"];
		}
		[ClassCleanup]
		public static void Cleanup()
		{
			_package.Dispose();
		}
        [TestMethod]
        public void VerifyCalculationPivotTable0()
        {
            var pt = _ptWs.PivotTables["PivotTable1"];
            pt.Calculate();

            Assert.AreEqual(86936.95, GetPtData(pt, 0, "Australia", "TRUE"));
            Assert.AreEqual(24.581, GetPtData(pt, 1, "Australia", "FALSE"));
            Assert.AreEqual(0.0134228187919463, (double)GetPtData(pt, 2, "Australia", null), 0.00000001D);

            Assert.AreEqual(335437D, GetPtData(pt, 0, "Peru", "true"));
            Assert.AreEqual(16.8274, GetPtData(pt, 1, "Peru", "true"));
            Assert.AreEqual(0.033557047, (double)GetPtData(pt, 2, "Peru", "true"), 0.0000001D);

            Assert.AreEqual(6529177.28, GetPtData(pt, 0, null, null));
            Assert.AreEqual(19.94585235, (double)GetPtData(pt, 1, null, null), 0.00000001D);
            Assert.AreEqual(1D, (double)GetPtData(pt, 2, null, null), 0.00000001D);
        }
        [TestMethod]
        public void VerifyCalculationPivotTable1()
        {
            var pt = _ptWs.PivotTables["PivotTable2"];
            pt.Calculate();

            Assert.AreEqual(49286.72, GetPtData(pt, 0, "Austria", "Niedersachsen", "TRUE"));
            Assert.AreEqual(54215D, GetPtData(pt, 1, "Austria", "Niedersachsen", "TRUE"));

            Assert.AreEqual(117336.43, GetPtData(pt, 0, "Belgium", null, "False"));
            Assert.AreEqual(129070D, GetPtData(pt, 1, "Belgium", null, "false"));

            Assert.AreEqual(8996331.09, GetPtData(pt, 0, null, null, null));
            Assert.AreEqual(9895964D, GetPtData(pt, 1, null, null, null));
        }
        [TestMethod]
        public void VerifyCalculationPivotTable2()
        {
            var pt = _ptWs.PivotTables["PivotTable3"];
            pt.Calculate();

            Assert.AreEqual(8996331.09, GetPtData(pt, 0));
            Assert.AreEqual(3727.881, GetPtData(pt, 1));
            Assert.AreEqual(9689.13, GetPtData(pt, 2));
            Assert.AreEqual(9895964.00, GetPtData(pt, 3));

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
