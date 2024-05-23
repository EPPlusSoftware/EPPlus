using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using Castle.DynamicProxy;
namespace EPPlusTest.Table.PivotTable.Calculation
{
    [TestClass]
    public class PivotTableCalculationSlicerTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelTable _tbl1, _tbl2;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableSlicer.xlsx", true);
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
        public void FilterSlicerSingleItem()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotWithSlicerSingle");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableSlicerSingle");
            pt.RowFields.Add(pt.Fields[0]);
			var slicer = pt.Fields[0].AddSlicer();
            slicer.SetPosition(1, 0, 8, 0);
            pt.CacheDefinition.Refresh();            
            var df = pt.DataFields.Add(pt.Fields["Price"]);

            foreach(var item in slicer.Cache.Data.Items)
            {
				item.Hidden= true;
			}

            slicer.Cache.Data.Items[1].Hidden = false;
            slicer.Cache.Data.Items.Refresh();
            ws.Cells["F5"].Formula = "GETPIVOTDATA(\"Price\",$C$3,\"Item\",\"Hammer\")";
			ws.Cells["F6"].Formula = "GETPIVOTDATA(\"Price\",$C$3,\"Item\",\"Crowbar\")";
			pt.Calculate();
            ws.Calculate();
            Assert.AreEqual(88.2D, ws.Cells["F5"].Value);
			Assert.AreEqual(ErrorValues.RefError, ws.Cells["F6"].Value);
		}
		[TestMethod]
		public void FilterSlicerMultipleItems()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotWithSlicerMultiple");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableSlicerSingle");
			pt.RowFields.Add(pt.Fields[0]);
			var slicer = pt.Fields[0].AddSlicer();
			slicer.SetPosition(1, 0, 8, 0);
			pt.CacheDefinition.Refresh();
			var df = pt.DataFields.Add(pt.Fields["Price"]);

			foreach (var item in slicer.Cache.Data.Items)
			{
				item.Hidden = true;
			}

			slicer.Cache.Data.Items[0].Hidden = false;
			slicer.Cache.Data.Items[1].Hidden = false;

			slicer.Cache.Data.Items.Refresh();

			pt.Calculate();
			ws.Cells["F5"].Formula = "GETPIVOTDATA(\"Price\",$C$3,\"Item\",\"Hammer\")";
			ws.Cells["F6"].Formula = "GETPIVOTDATA(\"Price\",$C$3,\"Item\",\"Crowbar\")";
			ws.Cells["F7"].Formula = "GETPIVOTDATA(\"Price\",$C$3,\"Item\",\"Saw\")";
			ws.Cells["F8"].Formula = "GETPIVOTDATA(\"Price\",$C$3)";

			ws.Calculate();
			Assert.AreEqual(88.2D, ws.Cells["F5"].Value);
			Assert.AreEqual(270.6D, ws.Cells["F6"].Value);
			Assert.AreEqual(ErrorValues.RefError, ws.Cells["F7"].Value);
			Assert.AreEqual(358.8, ws.Cells["F8"].Value);
		}
	}
}