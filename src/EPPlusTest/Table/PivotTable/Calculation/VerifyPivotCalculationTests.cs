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
		public void VerifyCalculationPivotTable0()
		{
			var pt = _ptWs.PivotTables[0];
			pt.Calculate();

			Assert.AreEqual(19979.5, GetPtData(pt, 0, "Jan", 2022));
			Assert.AreEqual(160368.28, GetPtData(pt, 0, "Jun", 2022));
			Assert.AreEqual(83965.85, GetPtData(pt, 0, "DEC", 2022));
			Assert.AreEqual(65971.35, GetPtData(pt, 0, "Mar", 2023));

			Assert.AreEqual(24125.35, GetPtData(pt, 1, "Jan", 2022));
			Assert.AreEqual(139010.99, GetPtData(pt, 1, "Jun", 2022));
			Assert.AreEqual(76737.86, GetPtData(pt, 1, "DEC", 2022));
			Assert.AreEqual(71281.65, GetPtData(pt, 1, "Mar", 2023));

			Assert.AreEqual(309812.948504, GetPtData(pt, 2, "Aug", 2022));
			Assert.AreEqual(0D, GetPtData(pt, 2, "Dec", 2023));

			Assert.AreEqual(243683.071401, GetPtData(pt, 3, "Oct", 2022));
			Assert.AreEqual(0D, GetPtData(pt, 2, "Dec", 2023));
		}
		[TestMethod]
		public void VerifyCalculationPivotTable1()
		{
			var pt = _ptWs.PivotTables[1];
			pt.Calculate();

			Assert.AreEqual(1743572.958631, GetPtData(pt, 0, "EP Catheters"));
			Assert.AreEqual(3867200.712841, GetPtData(pt, 0, null));
		}

		[TestMethod]
		public void VerifyCalculationPivotTable2()
		{
			var pt = _ptWs.PivotTables[2];
			pt.Calculate();

			Assert.AreEqual(2796.65, GetPtData(pt, 0, 2022, 2022, "Qtr2", "May", "Patient Care"));
			Assert.AreEqual(54706.29, GetPtData(pt, 0, 2022, 2022, "Qtr1", "Feb", null));
			Assert.AreEqual(1090330.25, GetPtData(pt, 0, null, null, null, null, null));
		}

		[TestMethod]
		public void VerifyCalculationPivotTable3()
		{
			var pt = _ptWs.PivotTables[3];
			pt.Calculate();

			Assert.AreEqual(97015.74, GetPtData(pt, 0, "jul", 2022));
			Assert.AreEqual(173647.98, GetPtData(pt, 0, null, 2023));
			Assert.AreEqual(1090330.25, GetPtData(pt, 0, null, null));
		}
		[TestMethod]
		public void VerifyCalculationPivotTable4()
		{
			var pt = _ptWs.PivotTables[4];
			pt.Calculate();

			Assert.AreEqual(3053395.883122, GetPtData(pt, 0, 2022));
			Assert.AreEqual(909570.888632003, (double)GetPtData(pt, 1, 2023), 0.00001D);
			Assert.AreEqual(4993589.981094, GetPtData(pt, 2, null));

		}
		[TestMethod]
		public void VerifyCalculationPivotTable5()
		{
			var pt = _ptWs.PivotTables[5];
			pt.Calculate();

			Assert.AreEqual(1038989.67, GetPtData(pt, 0, null, null));
			Assert.AreEqual(75067.77, GetPtData(pt, 0, "nov", 2022));
			Assert.AreEqual(0D, GetPtData(pt, 0, "nov", 2023));
			Assert.AreEqual(172691.67, GetPtData(pt, 0, null, 2023));
			Assert.AreEqual(30069.49, GetPtData(pt, 0, "oct", null));
		}
		[TestMethod]
		public void VerifyCalculationPivotTable6()
		{
			var pt = _ptWs.PivotTables[6];
			pt.Calculate();

			Assert.AreEqual(333214.299831, GetPtData(pt, 0, "Tissue Sealer"));
			Assert.AreEqual(1520214.663474, GetPtData(pt, 0, "EP Catheters"));
			Assert.AreEqual(43695.822963, (double)GetPtData(pt, 0, "Ultrasonic Scalpels"), 0.00001D);
		}
		[TestMethod]
		public void VerifyCalculationPivotTable7()
		{
			var pt = _ptWs.PivotTables[7];
			pt.Calculate();

			Assert.AreEqual(6141.33, GetPtData(pt, 0, "Suture Ret/Pssr"));
			Assert.AreEqual(523786.64, GetPtData(pt, 0, "EP Catheters"));
		}
		[TestMethod]
		public void VerifyCalculationPivotTable8()
		{
			var pt = _ptWs.PivotTables[8];
			pt.Calculate();

			Assert.AreEqual(108288.2, GetPtData(pt, 0, "Transseptal needles"));
		}
		[TestMethod]
		public void VerifyCalculationPivotTable9()
		{
			var pt = _ptWs.PivotTables[9];
			pt.Calculate();

			Assert.AreEqual(630648.87, GetPtData(pt, 0, "EP"));
		}

		[TestMethod]
		public void VerifyCalculationPivotTable10()
		{
			var pt = _ptWs.PivotTables[9];
			pt.Calculate();

			Assert.AreEqual(630648.87, GetPtData(pt, 0, "EP"));

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
