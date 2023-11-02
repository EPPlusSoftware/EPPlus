using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
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
        public void GetPivotData_Sum_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("SumRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable1");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.DataFields.Add(pt.Fields["Sales"]);
            pt.Refresh();
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(896D, ws.Cells["G5"].Value);
            Assert.AreEqual(818D, ws.Cells["G6"].Value);
            Assert.AreEqual(3188D, ws.Cells["G7"].Value);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Count_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("CountRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable1");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Count;
            pt.Refresh();
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(3D, ws.Cells["G5"].Value);
            Assert.AreEqual(6D, ws.Cells["G6"].Value);
            Assert.AreEqual(16D, ws.Cells["G7"].Value);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }

    }
}
