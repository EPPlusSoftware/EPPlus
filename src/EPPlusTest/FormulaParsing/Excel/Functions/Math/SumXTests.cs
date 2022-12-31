using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class SumXTests
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


        [TestMethod]
        public void SumX2My2_TwoRanges()
        {
            _sheet.Cells["A1"].Value = 5;
            _sheet.Cells["A2"].Value = 6;
            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["B1"].Value = 3;
            _sheet.Cells["B2"].Value = 4;
            _sheet.Cells["B3"].Value = 2;
            _sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,B1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(81d, _sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void SumX2My2_RangeAndArray()
        {
            _sheet.Cells["A1"].Value = 5;
            _sheet.Cells["A2"].Value = 6;
            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,{3;4;2})";
            _sheet.Calculate();
            Assert.AreEqual(81d, _sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void SumX2My2_NonMatchingLengths()
        {
            _sheet.Cells["A1"].Value = 5;
            _sheet.Cells["A2"].Value = 6;
            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["B1"].Value = 3;
            _sheet.Cells["B2"].Value = 4;
            _sheet.Cells["B3"].Value = 2;
            _sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,B1:B2)";
            _sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), _sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void SumX2My2_NonNumeric1()
        {
            _sheet.Cells["A1"].Value = 5;
            _sheet.Cells["A2"].Value = "6";
            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["B1"].Value = 3;
            _sheet.Cells["B2"].Value = 4;
            _sheet.Cells["B3"].Value = 2;
            _sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,B1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(61d, _sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void SumX2My2_NonNumeric2()
        {
            _sheet.Cells["A1"].Value = 5;
            _sheet.Cells["A2"].Value = "6";
            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["B1"].Value = 3;
            _sheet.Cells["B2"].Value = 4;
            _sheet.Cells["B3"].Value = "2";
            _sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,B1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(16d, _sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void SumXmY2_TwoRanges()
        {
            _sheet.Cells["A1"].Value = 5;
            _sheet.Cells["A2"].Value = 6;
            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["B1"].Value = 3;
            _sheet.Cells["B2"].Value = 4;
            _sheet.Cells["B3"].Value = 2;
            _sheet.Cells["C1"].Formula = "SUMXMY2(A1:A3,B1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(33d, _sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void SumX2pY2_TwoRanges()
        {
            _sheet.Cells["A1"].Value = 5;
            _sheet.Cells["A2"].Value = 6;
            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["B1"].Value = 3;
            _sheet.Cells["B2"].Value = 4;
            _sheet.Cells["B3"].Value = 2;
            _sheet.Cells["C1"].Formula = "SUMX2PY2(A1:A3,B1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(139d, _sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void SeriesSumShouldReturnCorrectResult()
        {
            _sheet.Cells["A1"].Formula = "SERIESSUM( 5, 1, 1, {1,1,1,1,1} )";

            _sheet.Cells["B1"].Value = 3;
            _sheet.Cells["B2"].Value = 4;
            _sheet.Cells["B3"].Value = 2;
            _sheet.Cells["C1"].Formula = "SeriesSum(2,1,1,B1:B3)";

            _sheet.Calculate();

            Assert.AreEqual(3905d, _sheet.Cells["A1"].Value, "First assert failed");
            Assert.AreEqual(38d, _sheet.Cells["C1"].Value, "Second assert failed");
        }
    }
}
