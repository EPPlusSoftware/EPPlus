using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class AggregateTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        private void LoadData1()
        {
            _sheet.Cells["A1"].Value = 3;
            _sheet.Cells["A2"].Value = 2.5;
            _sheet.Cells["A3"].Value = 1;
            _sheet.Cells["A4"].Value = 6;
            _sheet.Cells["A5"].Value = -2;
        }

        // Tests for Ignore nothing

        [TestMethod]
        public void ShouldHandleAverage()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 1, 4, A1, A2, A3, A4, A5 )";
            _sheet.Calculate();
            Assert.AreEqual(2.1d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void ShouldHandleSum()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 9, 4, A1, A2, A3, A4, A5 )";
            _sheet.Calculate();
            Assert.AreEqual(10.5D, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void ShouldHandleMin()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 5, 4, A1:A5 )";
            _sheet.Calculate();
            Assert.AreEqual(-2d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void ShouldHandleLarge()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 14, 4, A1:A5, 2 )";
            _sheet.Calculate();
            Assert.AreEqual(3d, _sheet.Cells["A6"].Value);
        }

        // Tests for Ignore hidden cells

        [TestMethod]
        public void HiddenCells_ShouldHandleAverage()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 1, 5, A1, A2, A3, A4, A5 )";
            _sheet.Row(3).Hidden = true;
            _sheet.Calculate();
            Assert.AreEqual(2.375d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void HiddenCells_ShouldHandleSum()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 9, 5, A1, A2, A3, A4, A5 )";
            _sheet.Row(3).Hidden = true;
            _sheet.Calculate();
            Assert.AreEqual(9.5d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void HiddenCells_ShouldHandleMin()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 5, 5, A1, A2, A3, A4, A5 )";
            _sheet.Row(5).Hidden = true;
            _sheet.Calculate();
            Assert.AreEqual(1d, _sheet.Cells["A6"].Value);
        }

        // Tests for ignoring errors

        [TestMethod]
        public void Errors_ShouldHandleAverage()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 1, 6, A1, A2, A3, A4, A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(2.375d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A3"].Value = 1;
            _sheet.Calculate();
            Assert.AreEqual(2.1d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleCount()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 2, 6, A1, A2, A3, A4, A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(4d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A3"].Value = 1;
            _sheet.Calculate();
            Assert.AreEqual(5d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleCountA()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 3, 6, A1, A2, A3, A4, A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(4d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A3"].Value = 1;
            _sheet.Calculate();
            Assert.AreEqual(5d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleMax()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 4, 6, A1, A2, A3, A4, A5 )";
            _sheet.Cells["A4"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(3d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A6"].Formula = "AGGREGATE( 4, 4, A1:A5 )";
            _sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleMin()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 5, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(-2d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A6"].Formula = "AGGREGATE( 5, 4, A1:A5 )";
            _sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleProduct()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 6, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(-90d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A6"].Formula = "AGGREGATE( 5, 4, A1:A5 )";
            _sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleStdevS()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 7, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            var result = (double)_sheet.Cells["A6"].Value;
            result = System.Math.Round(result, 5);
            Assert.AreEqual(3.30088d, result);

            _sheet.Cells["A6"].Formula = "AGGREGATE( 7, 4, A1:A5 )";
            _sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleStdevP()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 8, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            var result = (double)_sheet.Cells["A6"].Value;
            result = System.Math.Round(result, 5);
            Assert.AreEqual(2.85865d, result);

            _sheet.Cells["A6"].Formula = "AGGREGATE( 8, 4, A1:A5 )";
            _sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), _sheet.Cells["A6"].Value);
        }


        [TestMethod]
        public void Errors_ShouldHandleSum()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 9, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(9.5d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A3"].Value = 1;
            _sheet.Calculate();
            Assert.AreEqual(10.5d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleVarS()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 10, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            var result = (double)_sheet.Cells["A6"].Value;
            result = System.Math.Round(result, 5);
            Assert.AreEqual(10.89583d, result);

            _sheet.Cells["A3"].Value = 1;
            _sheet.Calculate();
            Assert.AreEqual(8.55d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleVarP()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 11, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(8.171875d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A3"].Value = 1;
            _sheet.Calculate();
            var result = (double)_sheet.Cells["A6"].Value;
            result = System.Math.Round(result, 2);
            Assert.AreEqual(6.84d, result);
        }

        [TestMethod]
        public void Errors_ShouldHandleMedian()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 12, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(2.75d, _sheet.Cells["A6"].Value);

            _sheet.Cells["A3"].Value = 1;
            _sheet.Calculate();
            var result = (double)_sheet.Cells["A6"].Value;
            result = System.Math.Round(result, 2);
            Assert.AreEqual(2.5d, result);
        }

        [TestMethod]
        public void Errors_ShouldHandleModeSngl()
        {
            LoadData1();
            _sheet.Cells["A2"].Value = 3;
            _sheet.Cells["A6"].Formula = "AGGREGATE( 13, 6, A1:A5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(3d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleLarge()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 14, 6, A1:A5, 1 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(6d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleSmall()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 15, 6, A1:A5, 1 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(-2d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandlePercentileInc()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 16, 6, A1:A5, 0 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(-2d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleQuartileInc()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 17, 6, A1:A5, 1 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(1.375d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandlePercentileExc()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 18, 6, A1:A5, 0.5 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(2.75d, _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void Errors_ShouldHandleQuartileExc()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 19, 6, A1:A5, 1 )";
            _sheet.Cells["A3"].Formula = "1/0";
            _sheet.Calculate();
            Assert.AreEqual(-0.875d, _sheet.Cells["A6"].Value);
        }

        // Tests for ignoring nested aggregate functions

        [TestMethod]
        public void IngoreNestedAggregateFunction()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 19, 6, A1:A5, 1)";
            _sheet.Cells["A7"].Formula = "AGGREGATE( 2, 0, A1:A6)";
            _sheet.Calculate();
            Assert.AreEqual(5d, _sheet.Cells["A7"].Value);
        }

        [TestMethod]
        public void IncludeNestedAggregateFunction()
        {
            LoadData1();
            _sheet.Cells["A6"].Formula = "AGGREGATE( 19, 6, A1:A5, 1)";
            _sheet.Cells["A7"].Formula = "AGGREGATE( 2, 4, A1:A6)";
            _sheet.Calculate();
            Assert.AreEqual(6d, _sheet.Cells["A7"].Value);
        }

        [TestMethod]
        public void ShouldHandleMultipleLevelsOfAggregate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet3 = package.Workbook.Worksheets.Add("sheet3");
                sheet3.Cells["A1"].Value = 26959.64;
                sheet3.Cells["A2"].Value = 82272d;
                sheet3.Cells["A3"].Formula = "AGGREGATE(9,0,A1:A2)";
                sheet3.Cells["A4"].Formula = "AGGREGATE(9,0,A1:A3)";

                var sheet2 = package.Workbook.Worksheets.Add("sheet2");
                sheet2.Cells["A1"].Formula = "sheet3!A4";
                package.Workbook.Calculate();
                Assert.AreEqual(109231.64d, sheet2.Cells["A1"].Value);

                sheet3.Cells["A3"].Formula = "AGGREGATE(8,0,A1:A2)";
                sheet3.Cells["A4"].Formula = "AGGREGATE(8,0,A1:A3)";
                package.Workbook.Calculate();
                Assert.AreEqual(27656.18, sheet2.Cells["A1"].Value);

                sheet3.Cells["A3"].Formula = "AGGREGATE(7,0,A1:A2)";
                sheet3.Cells["A4"].Formula = "AGGREGATE(7,0,A1:A3)";
                package.Workbook.Calculate();
                Assert.AreEqual(39111.7448d, System.Math.Round((double)sheet2.Cells["A1"].Value, 4));
            }
        }
    }
}
