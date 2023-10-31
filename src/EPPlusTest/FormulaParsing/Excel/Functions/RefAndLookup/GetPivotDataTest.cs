using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class GetPivotDataTest : TestBase
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _package;
        [TestInitialize]
        public void TestInitialize()
        {
            _package = OpenPackage("GetPivotDataTests.xlsx", true);
            _sheet = _package.Workbook.Worksheets.Add("Data");
            LoadHierarkiTestData(_sheet);
        }

        [TestCleanup]
        public void TestCleanup()
        {
            SaveAndCleanup(_package);
            _package.Dispose();
        }
        [TestMethod]
        public void GetPivotDataFromRowItem()
        {
            var ws = _package.Workbook.Worksheets.Add("RowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable1");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.DataFields.Add(pt.Fields["Sales"]);
            pt.CacheDefinition.Refresh();
            ws.Cells["E5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Calculate();

            Assert.AreEqual(896D, ws.Cells["E5"].Value);
        }
     }
}
