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
namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class GetPivotDataTest : TestBase
    {
        private static ExcelWorksheet _sheet;
        private static ExcelPackage _package;
        [ClassInitialize]
        public static void TestInitialize(TestContext context)
        {
            _package = OpenPackage("GetPivotDataTests.xlsx", true);
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
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.DataFields.Add(pt.Fields["Sales"]);
            pt.Calculate(true);
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
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_PercentOfGrandTotal()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_GrandTotal");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable11");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            pt.CacheDefinition.Refresh();
            df.ShowDataAs.SetPercentOfTotal();
            pt.Calculate();
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();
            Assert.AreEqual(0.281053952321205, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(0.256587202007528, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(1D, ws.Cells["G7"].Value);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_PercentOfColumnTotal()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_ColumnTotal");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable12");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentOfColumn();
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();
            Assert.AreEqual(0.790820829655781, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(1D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(1D, ws.Cells["G7"].Value);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_PercentOfRowTotal()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_RowTotal");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable13");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentOfRow();
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";
            ws.Calculate();
            Assert.AreEqual(1D, (double)ws.Cells["G5"].Value);
            Assert.AreEqual(0.256587202, (double)ws.Cells["G6"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["G7"].Value);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_RowAndCol_PercentOf()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_PercentOf_RC");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable14");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);

            var df1 = pt.DataFields.Add(pt.Fields["Sales"]);
            var df2 = pt.DataFields.Add(pt.Fields["Sales"]);
            var df3 = pt.DataFields.Add(pt.Fields["Sales"]);
            pt.CacheDefinition.Refresh();
            df1.Function = DataFieldFunctions.Sum;
            df1.ShowDataAs.SetPercent(pt.RowFields[0], pt.RowFields[0].Items.GetIndexByValue("USA"));
            df2.Function = DataFieldFunctions.Sum;
            df2.ShowDataAs.SetPercent(pt.RowFields[0], ePrevNextPivotItem.Previous);
            df3.Function = DataFieldFunctions.Sum;
            df3.ShowDataAs.SetPercent(pt.RowFields[0], ePrevNextPivotItem.Next);

            pt.Calculate();
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Canada\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Invalid Field\",\"North America\")";

            ws.Cells["H5"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Continent\",\"North America\",\"Country\",\"Canada\")";
            ws.Cells["H6"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["H7"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1)";
            ws.Cells["H8"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["H9"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Invalid Field\",\"North America\")";

            ws.Cells["I5"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Continent\",\"North America\",\"Country\",\"Canada\")";
            ws.Cells["I6"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["I7"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1)";
            ws.Cells["I8"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["I9"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Invalid Field\",\"North America\")";

            ws.Calculate();

            Assert.AreEqual(0.264508929, (double)ws.Cells["G5"].Value, 0.0000001D);
            Assert.AreEqual(0.694196429, (double)ws.Cells["G6"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["G7"].Value);
            Assert.AreEqual(ErrorValues.NullError, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G9"].Value);

            Assert.AreEqual(0D, (double)ws.Cells["H5"].Value, 0.0000001D);
            Assert.AreEqual(1.517073171, (double)ws.Cells["H6"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["H7"].Value);
            Assert.AreEqual(0D, ws.Cells["H8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["H9"].Value);

            Assert.AreEqual(0D, (double)ws.Cells["I5"].Value, 0.0000001D);
            Assert.AreEqual(1.94984326, (double)ws.Cells["I6"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["I7"].Value);
            Assert.AreEqual(ErrorValues.NullError, ws.Cells["I8"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["I9"].Value);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_Column_PercentOf()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_PercentOf_C");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable15");
            pt.ColumnFields.Add(pt.Fields["Country"]);

            var df1 = pt.DataFields.Add(pt.Fields["Sales"]);
            var df2 = pt.DataFields.Add(pt.Fields["Sales"]);
            var df3 = pt.DataFields.Add(pt.Fields["Sales"]);
            pt.CacheDefinition.Refresh();
            df1.Function = DataFieldFunctions.Sum;
            df1.ShowDataAs.SetPercent(pt.ColumnFields[0], pt.ColumnFields[0].Items.GetIndexByValue("USA"));
            df2.Function = DataFieldFunctions.Sum;
            df2.ShowDataAs.SetPercent(pt.ColumnFields[0], ePrevNextPivotItem.Previous);
            df3.Function = DataFieldFunctions.Sum;
            df3.ShowDataAs.SetPercent(pt.ColumnFields[0], ePrevNextPivotItem.Next);

            pt.Calculate();
            ws.Cells["G15"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Canada\")";
            ws.Cells["G16"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["G17"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"USA\")";
            ws.Cells["G18"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Sweden\")";
            ws.Cells["G19"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";

            ws.Cells["H15"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"Canada\")";
            ws.Cells["H16"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["H17"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"USA\")";
            ws.Cells["H18"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"Sweden\")";
            ws.Cells["H19"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1)";

            ws.Cells["I15"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"Canada\")";
            ws.Cells["I16"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["I17"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"USA\")";
            ws.Cells["I18"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"Sweden\")";
            ws.Cells["I19"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1)";

            ws.Calculate();

            Assert.AreEqual(0.264508929, (double)ws.Cells["G15"].Value, 0.0000001D);
            Assert.AreEqual(0.694196429, (double)ws.Cells["G16"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["G17"].Value, 0.0000001D);
            Assert.AreEqual(0.208705357, (double)ws.Cells["G18"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["G19"].Value, 0.0000001D);

            Assert.AreEqual(0.742946708, (double)ws.Cells["H15"].Value, 0.0000001D);
            Assert.AreEqual(1.517073171, (double)ws.Cells["H16"].Value, 0.0000001D);
            Assert.AreEqual(2.871794872, (double)ws.Cells["H17"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["H18"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["H19"].Value, 0.0000001D);

            Assert.AreEqual(1.156097561, (double)ws.Cells["I15"].Value, 0.0000001D);
            Assert.AreEqual(1.94984326, (double)ws.Cells["I16"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["I17"].Value, 0.0000001D);
            Assert.AreEqual(0.456097561, (double)ws.Cells["I18"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["I19"].Value, 0.0000001D);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_Row_PercentOf()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_PercentOf_R");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable16");
            pt.RowFields.Add(pt.Fields["Country"]);

            var df1 = pt.DataFields.Add(pt.Fields["Sales"]);
            var df2 = pt.DataFields.Add(pt.Fields["Sales"]);
            var df3 = pt.DataFields.Add(pt.Fields["Sales"]);
            pt.CacheDefinition.Refresh();
            df1.Function = DataFieldFunctions.Sum;
            df1.ShowDataAs.SetPercent(pt.RowFields[0], pt.RowFields[0].Items.GetIndexByValue("USA"));
            df2.Function = DataFieldFunctions.Sum;
            df2.ShowDataAs.SetPercent(pt.RowFields[0], ePrevNextPivotItem.Previous);
            df3.Function = DataFieldFunctions.Sum;
            df3.ShowDataAs.SetPercent(pt.RowFields[0], ePrevNextPivotItem.Next);

            pt.Calculate();
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Canada\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"USA\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";

            ws.Cells["H5"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"Canada\")";
            ws.Cells["H6"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["H7"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"USA\")";
            ws.Cells["H8"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"Sweden\")";
            ws.Cells["H9"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1)";

            ws.Cells["I5"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"Canada\")";
            ws.Cells["I6"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["I7"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"USA\")";
            ws.Cells["I8"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"Sweden\")";
            ws.Cells["I9"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1)";

            ws.Calculate();

            Assert.AreEqual(0.264508929, (double)ws.Cells["G5"].Value, 0.0000001D);
            Assert.AreEqual(0.694196429, (double)ws.Cells["G6"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["G7"].Value, 0.0000001D);
            Assert.AreEqual(0.208705357, (double)ws.Cells["G8"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["G9"].Value, 0.0000001D);

            Assert.AreEqual(0.742946708, (double)ws.Cells["H5"].Value, 0.0000001D);
            Assert.AreEqual(1.517073171, (double)ws.Cells["H6"].Value, 0.0000001D);
            Assert.AreEqual(2.871794872, (double)ws.Cells["H7"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["H8"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["H9"].Value, 0.0000001D);

            Assert.AreEqual(1.156097561, (double)ws.Cells["I5"].Value, 0.0000001D);
            Assert.AreEqual(1.94984326, (double)ws.Cells["I6"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["I7"].Value, 0.0000001D);
            Assert.AreEqual(0.456097561, (double)ws.Cells["I8"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["I9"].Value, 0.0000001D);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_RowAndCol_TopRef_PercentOf()
        {
            //TODO: Fix Asserts
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_PercentOf_RC_TR");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable17");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);

            var df1 = pt.DataFields.Add(pt.Fields["Sales"]);
            df1.Name = "Sales";
            var df2 = pt.DataFields.Add(pt.Fields["Sales"]);
            var df3 = pt.DataFields.Add(pt.Fields["Sales"]);
            pt.CacheDefinition.Refresh();
            df1.Function = DataFieldFunctions.Sum;
            df1.ShowDataAs.SetPercent(pt.RowFields[0], pt.RowFields[0].Items.GetIndexByValue("USA"));
            df2.Function = DataFieldFunctions.Sum;
            df2.ShowDataAs.SetPercent(pt.RowFields[0], ePrevNextPivotItem.Previous);
            df3.Function = DataFieldFunctions.Sum;
            df3.ShowDataAs.SetPercent(pt.RowFields[0], ePrevNextPivotItem.Next);

            pt.Calculate();
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Canada\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1)";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"Country\",\"Sweden\",\"State\",\"Stockholm\")";

            ws.Cells["H5"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Continent\",\"North America\",\"Country\",\"Canada\")";
            ws.Cells["H6"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["H7"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1)";
            ws.Cells["H8"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["H9"].Formula = "GETPIVOTDATA(\"Sales_2\",$A$1,\"Continent\",\"Europe\",\"Country\",\"Sweden\",\"State\",\"Stockholm\")";

            ws.Cells["I5"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Continent\",\"North America\",\"Country\",\"Canada\")";
            ws.Cells["I6"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Country\",\"Japan\")";
            ws.Cells["I7"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1)";
            ws.Cells["I8"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["I9"].Formula = "GETPIVOTDATA(\"Sales_3\",$A$1,\"Continent\",\"Europe\",\"Country\",\"Sweden\",\"State\",\"Stockholm\")";

            ws.Calculate();

            Assert.AreEqual(0.264508929, (double)ws.Cells["G5"].Value, 0.0000001D);
            Assert.AreEqual(0.694196429, (double)ws.Cells["G6"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["G7"].Value);
            Assert.AreEqual(ErrorValues.NullError, ws.Cells["G8"].Value);
            Assert.AreEqual(ErrorValues.NAError, ws.Cells["G9"].Value);

            Assert.AreEqual(1D, (double)ws.Cells["H5"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["H6"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["H7"].Value);  
            Assert.AreEqual(0D, ws.Cells["H8"].Value);
            Assert.AreEqual(1D, ws.Cells["H9"].Value);

            Assert.AreEqual(1D, (double)ws.Cells["I5"].Value, 0.0000001D);
            Assert.AreEqual(1D, (double)ws.Cells["I6"].Value, 0.0000001D);
            Assert.AreEqual(0D, (double)ws.Cells["H7"].Value);
            Assert.AreEqual(0D, (double)ws.Cells["I8"].Value);
            Assert.AreEqual(1D, (double)ws.Cells["I9"].Value);
        }

		[TestMethod]
        public void GetPivotData_Sum_ShowValueAs_PercentOfPartentRowTotal()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_ParentRowTotal");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable18");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentParentRow();
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Stockholm\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"Boston\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
            ws.Calculate();
            Assert.AreEqual(0.790820829655781, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(0.823529412, (double)ws.Cells["G6"].Value, 0.0000001);

            Assert.AreEqual(0.172991071, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(0.65830721, (double)ws.Cells["G9"].Value, 0.0000001);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_PercentOfPartentColumnTotal()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_ParentColTotal");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable19");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentParentColumn();
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Stockholm\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"Boston\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();
            Assert.AreEqual(1D, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(1D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(1D, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(1D, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0.256587202, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0.388017566, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(0.355395232, (double)ws.Cells["G12"].Value, 0.0000001);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_RunningTotal()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_RunningTotal");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable20");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetRunningTotal(pt.Fields["Country"]);
            pt.Calculate(true);

            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransico\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(1133D, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(33D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(411D, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(210D, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(0, (double)ws.Cells["G12"].Value, 0.0000001);
        }
		[TestMethod]
		public void GetPivotData_Sum_ShowValueAs_PercentOfRunningTotal()
		{
			var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_PercentRunningTotal");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable21");
			pt.ColumnFields.Add(pt.Fields["Continent"]);
			pt.RowFields.Add(pt.Fields["Country"]);
			pt.RowFields.Add(pt.Fields["State"]);
			var df = pt.DataFields.Add(pt.Fields["Sales"]);
			df.Function = DataFieldFunctions.Sum;
			df.ShowDataAs.SetPercentOfRunningTotal(pt.Fields["Country"]);
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransico\")";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
			ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
			ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
			ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
			ws.Calculate();

			Assert.AreEqual(1D, (double)ws.Cells["G5"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G6"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G7"].Value, 0.0000001);
			Assert.AreEqual(0D, ws.Cells["G8"].Value);
			Assert.AreEqual(1D, (double)ws.Cells["G9"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G11"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G12"].Value, 0.0000001);
		}
		[TestMethod]
		public void GetPivotData_Sum_ShowValueAs_RankAscending()
		{
			var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_RankAscending");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable22");
			pt.ColumnFields.Add(pt.Fields["Continent"]);
			pt.RowFields.Add(pt.Fields["Country"]);
			pt.RowFields.Add(pt.Fields["State"]);
			var df = pt.DataFields.Add(pt.Fields["Sales"]);
			df.Function = DataFieldFunctions.Sum;
			df.ShowDataAs.SetRankAscending(pt.Fields["Country"]);
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransico\")";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
			ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
			ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
			ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Cells["G13"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Vietnam\")";
			ws.Cells["G14"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Japan\")";
			ws.Cells["G15"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"France\")";
			ws.Cells["G16"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\", \"Europe\", \"Country\", \"Sweden\")";

			ws.Calculate();

			Assert.AreEqual(2D, (double)ws.Cells["G5"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G6"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G7"].Value, 0.0000001);
			Assert.AreEqual(0D, ws.Cells["G8"].Value);
			Assert.AreEqual(1D, (double)ws.Cells["G9"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G11"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G12"].Value, 0.0000001);

            Assert.AreEqual(6D, (double)ws.Cells["G13"].Value, 0.0000001);
			Assert.AreEqual(7D, (double)ws.Cells["G14"].Value, 0.0000001);
			Assert.AreEqual(4D, (double)ws.Cells["G15"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G16"].Value, 0.0000001);
		}
		[TestMethod]
		public void GetPivotData_Sum_ShowValueAs_RankDescending()
		{
			var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_RankDescending");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable23");
			pt.ColumnFields.Add(pt.Fields["Continent"]);
			pt.RowFields.Add(pt.Fields["Country"]);
			pt.RowFields.Add(pt.Fields["State"]);
			var df = pt.DataFields.Add(pt.Fields["Sales"]);
			df.Function = DataFieldFunctions.Sum;
			df.ShowDataAs.SetRankDescending(pt.Fields["Country"]);
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransico\")";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
			ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
			ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
			ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
			ws.Cells["G13"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Vietnam\")";
			ws.Cells["G14"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"Japan\")";
			ws.Cells["G15"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Country\",\"France\")";
            ws.Cells["G16"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\", \"Europe\", \"Country\", \"Sweden\")";

			ws.Calculate();

			Assert.AreEqual(1D, (double)ws.Cells["G5"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G6"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G7"].Value, 0.0000001);
			Assert.AreEqual(0D, ws.Cells["G8"].Value);
			Assert.AreEqual(1D, (double)ws.Cells["G9"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G11"].Value, 0.0000001);
			Assert.AreEqual(0D, (double)ws.Cells["G12"].Value, 0.0000001);

			Assert.AreEqual(3D, (double)ws.Cells["G13"].Value, 0.0000001);
			Assert.AreEqual(2D, (double)ws.Cells["G14"].Value, 0.0000001);
			Assert.AreEqual(5D, (double)ws.Cells["G15"].Value, 0.0000001);
			Assert.AreEqual(3D, (double)ws.Cells["G16"].Value, 0.0000001);
		}



		[TestMethod]
        public void PivotItemKeyTests()
        {
            //The pivot item key is used for aggregating items per row/column fields.
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0 }, 1));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0 }, 1));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue }, 1));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 1));

            //2 row and 1 col
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0, 0 }, 2));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0, PivotCalculationStore.SumLevelValue }, 2));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue, 0 }, 2));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 2));            
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue, 0 }, 2));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 2));
            
            Assert.IsTrue(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0, PivotCalculationStore.SumLevelValue }, 2));
            Assert.IsTrue(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0, 0 }, 2));

            //1 row and 2 col
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0, 0 }, 1));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0, PivotCalculationStore.SumLevelValue }, 1));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 1));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0, PivotCalculationStore.SumLevelValue }, 1));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0, 0 }, 1));
            Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 1));

            Assert.IsTrue(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue, 0 }, 1));
            Assert.IsTrue(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue, 0 }, 1));
        }
    }
}
