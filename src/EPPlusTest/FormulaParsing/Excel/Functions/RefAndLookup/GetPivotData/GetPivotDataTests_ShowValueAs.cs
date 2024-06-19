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
using System.Globalization;
namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class GetPivotDataTests_ShowValueAs : TestBase
    {
        private static ExcelWorksheet _sheet;
        private static ExcelPackage _package;
        [ClassInitialize]
        public static void TestInitialize(TestContext context)
        {
            _package = OpenPackage("GetPivotData_ShowValueAs.xlsx", true);
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
        public void GetPivotData_Sum_ShowValueAs_PercentParentTotal()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_ParentTotal");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable13");
            pt.RowFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentParent(pt.RowFields[1]);
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();
            Assert.AreEqual(1D, (double)ws.Cells["G5"].Value, 0.0000001);
            
            Assert.AreEqual(0.176470588, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(0.458705357, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["G8"].Value);
            Assert.AreEqual(0.65830721, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G12"].Value, 0.0000001);
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
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_RunningTot-c");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable20");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetRunningTotal(pt.Fields["Country"]);
            pt.Calculate(true);

            var current = CultureInfo.CurrentCulture;

            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
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
        public void GetPivotData_Sum_ShowValueAs_RunningTotal_column_Top()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_RunTot_c");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable20-c");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetRunningTotal(pt.Fields["Continent"]);
            pt.Calculate(true);

            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(896D, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(33D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(411D, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(187D, ws.Cells["G8"].Value);
            Assert.AreEqual(0D, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(818D, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(2055D, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(3188D, (double)ws.Cells["G12"].Value, 0.0000001);
        }


        [TestMethod]
		public void GetPivotData_Sum_ShowValueAs_PercentOfRunningTotal_row_top()
		{
			var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_PercentRunTot-r");
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
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
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
        public void GetPivotData_Sum_ShowValueAs_PercentOfRunningTotal_column_Top()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_PercentRunTot_c");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable21-c");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentOfRunningTotal(pt.Fields["Continent"]);
            pt.Calculate(true);

            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(1D, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(1D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(1D, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(1D, ws.Cells["G8"].Value);
            Assert.AreEqual(0D, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0.256587202, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0.644604768, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(1D, (double)ws.Cells["G12"].Value, 0.0000001);
        }

        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_PercentOfRunningTotal_Scattered()
        {
            using(var p = OpenPackage("PercentOfRunningTotal_Scattered.xlsx", true))
            {
                var dataSheet = p.Workbook.Worksheets.Add("datasheet");
                var ptSheet = p.Workbook.Worksheets.Add("pivotData");

                var dataTable = dataSheet.Tables.Add(dataSheet.Cells["A1:C6"], "dataTable");

                dataSheet.Cells["A1"].Value = "Continent";
                dataSheet.Cells["A2"].Value = "Europe";
                dataSheet.Cells["A3"].Value = "Europe";
                dataSheet.Cells["A4"].Value = "South America";
                dataSheet.Cells["A5"].Value = "Asia";
                dataSheet.Cells["A6"].Value = "Africa";

                dataSheet.Cells["B1"].Value = "Country";
                dataSheet.Cells["B2"].Value = "Sweden";
                dataSheet.Cells["B3"].Value = "Spain";
                dataSheet.Cells["B4"].Value = "Brazil";
                dataSheet.Cells["B5"].Value = "Japan";
                dataSheet.Cells["B6"].Value = "Uganda";

                dataSheet.Cells["C1"].Value = "Sales";
                dataSheet.Cells["C2"].Value = 120;
                dataSheet.Cells["C3"].Value = 300;
                dataSheet.Cells["C4"].Value = 90;
                dataSheet.Cells["C5"].Value = 600;
                dataSheet.Cells["C6"].Value = 3500;

                var pt = ptSheet.PivotTables.Add(ptSheet.Cells["A1"], dataTable, "PercentOfRunningTotal");

                pt.RowFields.Add(pt.Fields["Continent"]);
                pt.RowFields.Add(pt.Fields["Country"]);

                var dt = pt.DataFields.Add(pt.Fields["Sales"]);

                pt.DataOnRows = false;

                dt.ShowDataAs.SetPercentOfRunningTotal(pt.Fields["Continent"]);

                ptSheet.Cells["A14"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
                ptSheet.Cells["A15"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"South America\")";
                ptSheet.Cells["A16"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
                ptSheet.Cells["A17"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Africa\")";

                pt.Calculate();
                ptSheet.Calculate();

                Assert.AreEqual(0.0911062D, (double)ptSheet.Cells["A14"].Value, 0.0000001);
                Assert.AreEqual(0.1106290D, (double)ptSheet.Cells["A15"].Value, 0.0000001);
                Assert.AreEqual(0.2407809D, (double)ptSheet.Cells["A16"].Value, 0.0000001);
                Assert.AreEqual(1D, (double)ptSheet.Cells["A17"].Value, 0.0000001);

                SaveAndCleanup(p);
            }
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
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
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
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
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
		public void GetPivotData_Sum_ShowValueAs_Index()
		{
			var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_Index");
			var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable24");
			pt.ColumnFields.Add(pt.Fields["Continent"]);
			pt.RowFields.Add(pt.Fields["Country"]);
			pt.RowFields.Add(pt.Fields["State"]);
			var df = pt.DataFields.Add(pt.Fields["Sales"]);
			df.Function = DataFieldFunctions.Sum;
			df.ShowDataAs.SetIndex();
			pt.Calculate(true);

			ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
			ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
			ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
			ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
			ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
			ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
			ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
			ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
			ws.Calculate();

			Assert.AreEqual(2.8137687555, (double)ws.Cells["G5"].Value, 0.0000001);
			Assert.AreEqual(3.89731051345, (double)ws.Cells["G6"].Value, 0.0000001);
			Assert.AreEqual(2.81376875552, (double)ws.Cells["G7"].Value, 0.0000001);
			Assert.AreEqual(0D, ws.Cells["G8"].Value);
			Assert.AreEqual(1D, (double)ws.Cells["G9"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G10"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G11"].Value, 0.0000001);
			Assert.AreEqual(1D, (double)ws.Cells["G12"].Value, 0.0000001);
		}
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_Difference_Previous()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_Difference_Prev");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable25");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetDifference(pt.RowFields[1], ePrevNextPivotItem.Previous);
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();
            Assert.AreEqual(0D, (double)ws.Cells["G5"].Value, 0.0000001);

            Assert.AreEqual(-121D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(81D, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(101D, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G12"].Value, 0.0000001);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_Difference_Next()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_Difference_Next");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable26");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetDifference(pt.RowFields[1], ePrevNextPivotItem.Next);
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Stockholm\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"New York\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Frankfurt\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();
            Assert.AreEqual(0D, (double)ws.Cells["G5"].Value, 0.0000001);

            Assert.AreEqual(121D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(-81D, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(-101D, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G12"].Value, 0.0000001);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_Difference_Continent_Europe()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_Difference_Item");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable27");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetDifference(pt.ColumnFields[0], 0);
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Stockholm\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"New York\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Frankfurt\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();
            Assert.AreEqual(896D, (double)ws.Cells["G5"].Value, 0.0000001);

            Assert.AreEqual(0D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(330D, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(-187D, ws.Cells["G8"].Value);
            Assert.AreEqual(0, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(419D, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(315D, (double)ws.Cells["G12"].Value, 0.0000001);        
        }

        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_DifferencePercent_Previous()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_DiffPerc_Prev");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable28");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentageDifference(pt.RowFields[1], ePrevNextPivotItem.Previous);
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Västerås\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"San Fransisco\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Berlin\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Cells["G13"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\",\"State\",\"Västerås\")";
            ws.Calculate();
            
            Assert.AreEqual(0D, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(-0.785714286, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(0.245454545, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(0.926605505, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G12"].Value, 0.0000001);
            Assert.AreEqual(ErrorValues.NullError, ws.Cells["G13"].Value);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_DifferencePercent_Next()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_DiffPerc_Next");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable29");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentageDifference(pt.RowFields[1], ePrevNextPivotItem.Next);
            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Stockholm\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"New York\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Frankfurt\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Cells["G13"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\",\"State\",\"Stockholm\")";
            ws.Calculate();
            Assert.AreEqual(0D, (double)ws.Cells["G5"].Value, 0.0000001);

            Assert.AreEqual(3.666666667D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(-0.197080292, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(0D, ws.Cells["G8"].Value);
            Assert.AreEqual(-0.480952381, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G12"].Value, 0.0000001);
            Assert.AreEqual(ErrorValues.NullError, ws.Cells["G13"].Value);
        }
        [TestMethod]
        public void GetPivotData_Sum_ShowValueAs_DifferencePercent_Continent_Europe()
        {
            var ws = _package.Workbook.Worksheets.Add("Sum_ShowDataAs_DiffPerc_Item");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _sheet.Cells["A1:D17"], "PivotTable31");
            pt.ColumnFields.Add(pt.Fields["Continent"]);
            pt.RowFields.Add(pt.Fields["Country"]);
            pt.RowFields.Add(pt.Fields["State"]);
            var df = pt.DataFields.Add(pt.Fields["Sales"]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentageDifference(pt.ColumnFields[0], 0);

            pt.Calculate(true);
            ws.Cells["G5"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\")";
            ws.Cells["G6"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\",\"State\",\"Stockholm\")";
            ws.Cells["G7"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"USA\",\"State\",\"New York\")";
            ws.Cells["G8"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\",\"Country\",\"Sweden\")";
            ws.Cells["G9"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"State\",\"Frankfurt\")";
            ws.Cells["G10"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Europe\")";
            ws.Cells["G11"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"Asia\")";
            ws.Cells["G12"].Formula = "GETPIVOTDATA(\"Sales\",$A$1,\"Continent\",\"North America\")";
            ws.Calculate();

            Assert.AreEqual(0D, (double)ws.Cells["G5"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G6"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G7"].Value, 0.0000001);
            Assert.AreEqual(ErrorValues.NullError, ws.Cells["G8"].Value);
            Assert.AreEqual(0, (double)ws.Cells["G9"].Value, 0.0000001);
            Assert.AreEqual(0D, (double)ws.Cells["G10"].Value, 0.0000001);
            Assert.AreEqual(0.512224939D, (double)ws.Cells["G11"].Value, 0.0000001);
            Assert.AreEqual(0.385085575, (double)ws.Cells["G12"].Value, 0.0000001);
        }   
    }
}
