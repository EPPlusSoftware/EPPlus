using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class VstackTests
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _package;

        [TestInitialize]
        public void TestInitialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void TestCleanup()
        {
            _package.Dispose();
        }

        private void AddTestData()
        {
            var n = 1;
            for (var col = 1; col < 4; col++)
            {
                for (var row = 1; row < 3; row++)
                {
                    _sheet.Cells[row, col].Value = n++;
                }
            }
        }

        [TestMethod]
        public void BasicTest1()
        {
            AddTestData();
            _sheet.Cells[1, 5].Value = 100;
            _sheet.Cells[2, 5].Value = 200;
            _sheet.Cells[10, 1].Formula = "VSTACK(A1:C2,E1:E2)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells[10, 1].Value);
            Assert.AreEqual(3, _sheet.Cells[10, 2].Value);
            Assert.AreEqual(5, _sheet.Cells[10, 3].Value);
            Assert.AreEqual(2, _sheet.Cells[11, 1].Value);
            Assert.AreEqual(4, _sheet.Cells[11, 2].Value);
            Assert.AreEqual(6, _sheet.Cells[11, 3].Value);
            Assert.AreEqual(100, _sheet.Cells[12, 1].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[12, 2].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[12, 3].Value);
            Assert.AreEqual(200, _sheet.Cells[13, 1].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[13, 2].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[13, 3].Value);
        }

        [TestMethod]
        public void VStackWithSingleArg()
        {
            AddTestData();
            _sheet.Cells[1, 5].Value = 100;
            _sheet.Cells[2, 5].Value = 200;
            _sheet.Cells[10, 1].Formula = "VSTACK(A1:C2,E1:E2,1000)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells[10, 1].Value);
            Assert.AreEqual(3, _sheet.Cells[10, 2].Value);
            Assert.AreEqual(5, _sheet.Cells[10, 3].Value);
            Assert.AreEqual(2, _sheet.Cells[11, 1].Value);
            Assert.AreEqual(4, _sheet.Cells[11, 2].Value);
            Assert.AreEqual(6, _sheet.Cells[11, 3].Value);
            Assert.AreEqual(100, _sheet.Cells[12, 1].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[12, 2].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[12, 3].Value);
            Assert.AreEqual(200, _sheet.Cells[13, 1].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[13, 2].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[13, 3].Value);
            Assert.AreEqual(1000d, _sheet.Cells[14, 1].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[14, 2].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[14, 3].Value);
        }
    }
}
