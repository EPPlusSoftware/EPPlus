using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Slicer
{
    [TestClass]
    public class SlicerCopyTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("SlicerCopy.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void CopyTableSlicer()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSlicerSource");

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table2");
            var slicer = ws.Drawings.AddTableSlicer(tbl.Columns[1]);
            slicer.SetPosition(1, 0, 5, 0);

            slicer.SetSize(200, 600);

            var copy = _pck.Workbook.Worksheets.Add("TableSlicerCopy", ws);
        }
        [TestMethod]
        public void CopyPivotTableSlicer()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotTableSlicerSource");

            LoadTestdata(ws);
            var pt = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Table3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);
            var slicer = ws.Drawings.AddPivotTableSlicer(pt.Fields[3]);
            slicer.SetPosition(1, 0, 8, 0);

            slicer.SetSize(200, 600);

            var copy = _pck.Workbook.Worksheets.Add("PivotTableSlicerCopy", ws);
        }
		[TestMethod]
		public void CopyPivotTableSlicerToExternalPackage()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotTableSlicerSourceExt");

			LoadTestdata(ws);
			var pt = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Table3");
			pt.RowFields.Add(pt.Fields[1]);
			pt.DataFields.Add(pt.Fields[3]);
			var slicer = ws.Drawings.AddPivotTableSlicer(pt.Fields[3]);
			slicer.SetPosition(1, 0, 8, 0);

			slicer.SetSize(200, 600);

			using(var p2=new ExcelPackage())
            {
				var copy = p2.Workbook.Worksheets.Add("PivotTableSlicerCopy", ws);
                Assert.AreEqual(8, copy.Drawings[0].From.Column);
				Assert.AreEqual(1, copy.Drawings[0].From.Row);
				SaveWorkbook("SlicerCopyNewWb.xlsx", p2);
			}
		}
	}
}
