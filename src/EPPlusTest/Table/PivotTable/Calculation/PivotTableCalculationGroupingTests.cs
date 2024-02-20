﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable.Calculation
{
	[TestClass]
	public class PivotTableCalculationGroupingTests : TestBase
	{
		static ExcelPackage _pck;
		static ExcelTable _tbl1, _tbl2;
		[ClassInitialize]
		public static void Init(TestContext context)
		{
			InitBase();
			_pck = OpenPackage("PivotTableCalculationGrouping.xlsx", true);
			var ws = _pck.Workbook.Worksheets.Add("Data1");
			var r = LoadItemData(ws);
			_tbl1 = ws.Tables.Add(r, "Table1");
			ws = _pck.Workbook.Worksheets.Add("Data2");
			r = LoadItemData(ws);
			_tbl2 = ws.Tables.Add(r, "Table2");
		}
		[ClassCleanup]
		public static void Cleanup()
		{
			SaveAndCleanup(_pck);
		}
		[TestMethod]
		public void DateGroupYearCalculation()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotYearGrouping");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableYearGrouping");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.AddDateGrouping(eDateGroupBy.Years);
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			pt.CacheDefinition.Refresh();
			

			pt.Calculate();
			Assert.AreEqual(445.52, pt.CalculatedItems[0][[0]]);
		}
		[TestMethod]
		public void DateGroupYearMonthCalculation()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotYearGrouping");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableYearGrouping");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			rf.AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months);
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			pt.CacheDefinition.Refresh();


			pt.Calculate();
			Assert.AreEqual(85.2, pt.CalculatedItems[0][[0, 0]]);
			Assert.AreEqual(12.2, pt.CalculatedItems[0][[0, 1]]);
			Assert.AreEqual(445.52, pt.CalculatedItems[0][[int.MaxValue, int.MaxValue]]);
		}

	}
}
