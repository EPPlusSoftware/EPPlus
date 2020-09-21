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

    }
}
