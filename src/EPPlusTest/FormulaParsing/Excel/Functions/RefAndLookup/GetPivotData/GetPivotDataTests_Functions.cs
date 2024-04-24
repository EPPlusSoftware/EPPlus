using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using OfficeOpenXml.Table.PivotTable.Calculation;
using FakeItEasy;
using OfficeOpenXml.ConditionalFormatting;
namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class GetPivotDataTests_Functions : TestBase
    {
        private static ExcelWorksheet _sheet;
        private static ExcelPackage _package;
        [ClassInitialize]
        public static void TestInitialize(TestContext context)
        {
            _package = OpenPackage("GetPivotData_Functions.xlsx", true);
            _sheet = _package.Workbook.Worksheets.Add("Data");
            LoadHierarkiTestData(_sheet);
        }

        [ClassCleanup]
        public static void TestCleanup()
        {
            SaveAndCleanup(_package);
            _package.Dispose();
        }
        [TestMethod]
        public void GetPivotData_Sum_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("SumRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable1");
            var columnField = pt.ColumnFields.Add(pt.Fields["Continent"]);
            var rowField = pt.RowFields.Add(pt.Fields["Country"]);
            pt.DataFields.Add(pt.Fields["Sales"]);
            pt.Calculate(true);
            pt.GetPivotData("Sales", new List<PivotDataFieldItemSelection> { new  PivotDataFieldItemSelection("Continent", "North America"), new PivotDataFieldItemSelection("Country", "USA") });
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
        public void GetPivotData_Sum_TwoRowField_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("SumRowData_2df");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable1");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["State"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.DataFields.Add(pt.Fields["Sales"]);
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G5"].Value);
            Assert.AreEqual(818D, ws.Cells["G6"].Value);
            Assert.AreEqual(3188D, ws.Cells["G7"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Count_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("CountRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable2");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Count;
            pt.Calculate(true);
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
        [TestMethod]
        public void GetPivotData_Min_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("MinRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable3");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Min;
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(155D, ws.Cells["G5"].Value);
            Assert.AreEqual(33D, ws.Cells["G6"].Value);
            Assert.AreEqual(33D, ws.Cells["G7"].Value);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Max_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("MaxRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable4");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Max;
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(411D, ws.Cells["G5"].Value);
            Assert.AreEqual(210D, ws.Cells["G6"].Value);
            Assert.AreEqual(534D, ws.Cells["G7"].Value);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Product_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("ProductRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable5");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Product;
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(21022650D, ws.Cells["G5"].Value);
            Assert.AreEqual(2733102395100D, ws.Cells["G6"].Value);
            Assert.AreEqual(2.14276220630102E+35D, (double)ws.Cells["G7"].Value, 1E20);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Average_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("AverageRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable6");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Average;
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(298.666666666667D, (double)ws.Cells["G5"].Value, 0.00001D);
            Assert.AreEqual(136.333333333333D, (double)ws.Cells["G6"].Value, 0.00001D);
            Assert.AreEqual(199.25, ws.Cells["G7"].Value);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_StdDev_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("StdevRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable7");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.StdDev;
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\",\"Country\",\"China\")";
            ws.Calculate();

            Assert.AreEqual(130.844691651337, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(62.640774792995d, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(134.901198413258d, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
            Assert.AreEqual(ErrorValues.Div0Error, ws.Cells["G10"].Value);
        }
        [TestMethod]
        public void GetPivotData_StdDevP_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("StdevPRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable8");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.StdDevP;
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\",\"Country\",\"China\")";
            ws.Calculate();

            Assert.AreEqual(106.834243365859, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(57.182942289540d, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(130.617523709493d, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
            Assert.AreEqual(0D, ws.Cells["G10"].Value);
        }
        [TestMethod]
        public void GetPivotData_Var_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("VarRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable9");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Var;
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\",\"Country\",\"China\")";
            ws.Calculate();

            Assert.AreEqual(17120.3333333333d, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(3923.86666666667d, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(18198.3333333333d, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
            Assert.AreEqual(ErrorValues.Div0Error, ws.Cells["G10"].Value);
        }
        [TestMethod]
        public void GetPivotData_VarP_DataField()
        {
            var ws = _package.Workbook.Worksheets.Add("VarPRowData");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable10");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.VarP;
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\",\"Country\",\"China\")";
            ws.Calculate();

            Assert.AreEqual(11413.5555555556d, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(3269.88888888889d, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(17060.9375d, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
            Assert.AreEqual(0D, ws.Cells["G10"].Value);
        }
	}
}
