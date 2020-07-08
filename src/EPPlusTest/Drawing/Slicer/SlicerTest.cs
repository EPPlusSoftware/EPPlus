using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicer;
using System.IO;

namespace EPPlusTest.Drawing.Slicer
{
    [TestClass]
    public class SlicerTest : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("SlicerText.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Richtext");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);

            File.Copy(fileName, dirName + "\\SlicerRead.xlsx", true);
        }
        [TestMethod]
        public void ReadSlicer()
        {
            using (var p = OpenTemplatePackage("Slicer.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual(5, ws.Drawings.Count);
                Assert.IsInstanceOfType(ws.Drawings[0], typeof(ExcelTableSlicer));
                Assert.AreEqual(1, ws.SlicerXmlSources._list.Count);

                var tableSlicer = ws.Drawings[0].As.Slicer.TableSlicer;
                Assert.AreEqual(eSlicerStyle.None, tableSlicer.Style);
                Assert.AreEqual("CompanyName", tableSlicer.Caption);
                Assert.AreEqual("CompanyName", tableSlicer.Name);
                Assert.AreEqual("Slicer_CompanyName", tableSlicer.CacheName);
                Assert.AreEqual(0, tableSlicer.StartItem);
                Assert.AreEqual(19, tableSlicer.RowHeight);
                Assert.AreEqual(1, tableSlicer.ColumnCount);
                Assert.IsNotNull(tableSlicer.Cache);
                Assert.AreEqual(1, tableSlicer.Cache.TableId);
                Assert.AreEqual(1, tableSlicer.Cache.ColumnId);
                Assert.IsNotNull(tableSlicer.Cache.TableColumn);
                
                ws = p.Workbook.Worksheets[1];
                Assert.AreEqual(4, ws.Drawings.Count);
                Assert.IsInstanceOfType(ws.Drawings[1], typeof(ExcelPivotTableSlicer));
                Assert.IsInstanceOfType(ws.Drawings[2], typeof(ExcelPivotTableSlicer));
                Assert.IsInstanceOfType(ws.Drawings[3], typeof(ExcelTableSlicer));
                Assert.AreEqual(2, ws.SlicerXmlSources._list.Count);

                var pivotTableslicer = ws.Drawings[1].As.Slicer.PivotTableSlicer;
                Assert.AreEqual(eSlicerStyle.None, pivotTableslicer.Style);
                Assert.AreEqual("CompanyName", pivotTableslicer.Caption);
                Assert.AreEqual("CompanyName 1", pivotTableslicer.Name);
                Assert.AreEqual("Slicer_CompanyName1", pivotTableslicer.CacheName);
                Assert.AreEqual(4, pivotTableslicer.StartItem);
                Assert.AreEqual(19, pivotTableslicer.RowHeight);
                Assert.AreEqual(1, pivotTableslicer.ColumnCount);
                Assert.AreEqual(1, pivotTableslicer.Cache.PivotTables.Count);
            }
        }
        [TestMethod]
        public void AddTableSlicer()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSlicer");

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"],"Table1");
            var slicer = ws.Drawings.AddTableSlicer("Slicer1", tbl.Columns[0]);
        }
    }
}
