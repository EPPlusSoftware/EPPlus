using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class SortTests
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

        [TestMethod]
        public void BasicByColTestAsc()
        {
            _sheet.Cells[1, 1].Value = 3;
            _sheet.Cells[2, 1].Value = 2;
            _sheet.Cells[3, 1].Value = 1;

            _sheet.Cells[1, 2].Value = "C";
            _sheet.Cells[2, 2].Value = "B";
            _sheet.Cells[3, 2].Value = "A";

            _sheet.Cells[4, 1].Formula = "SORT(A1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(3, _sheet.Cells[6, 1].Value);
            Assert.AreEqual("A", _sheet.Cells[4, 2].Value);
            Assert.AreEqual("B", _sheet.Cells[5, 2].Value);
            Assert.AreEqual("C", _sheet.Cells[6, 2].Value);
        }

        [TestMethod]
        public void BasicByColTest_ColIx1_Asc()
        {
            _sheet.Cells[1, 1].Value = 3;
            _sheet.Cells[2, 1].Value = 2;
            _sheet.Cells[3, 1].Value = 1;

            _sheet.Cells[1, 2].Value = "B";
            _sheet.Cells[2, 2].Value = "C";
            _sheet.Cells[3, 2].Value = "A";

            _sheet.Cells[4, 1].Formula = "SORT(A1:B3, 2)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(3, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[6, 1].Value);
            Assert.AreEqual("A", _sheet.Cells[4, 2].Value);
            Assert.AreEqual("B", _sheet.Cells[5, 2].Value);
            Assert.AreEqual("C", _sheet.Cells[6, 2].Value);
        }

        [TestMethod]
        public void BasicByColTestDesc()
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[2, 1].Value = 2;
            _sheet.Cells[3, 1].Value = 3;

            _sheet.Cells[1, 2].Value = "A";
            _sheet.Cells[2, 2].Value = "B";
            _sheet.Cells[3, 2].Value = "C";

            _sheet.Cells[4, 1].Formula = "SORT(A1:B3, 1, -1)";
            _sheet.Calculate();
            Assert.AreEqual(3, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(1, _sheet.Cells[6, 1].Value);
            Assert.AreEqual("C", _sheet.Cells[4, 2].Value);
            Assert.AreEqual("B", _sheet.Cells[5, 2].Value);
            Assert.AreEqual("A", _sheet.Cells[6, 2].Value);
        }

        [TestMethod]
        public void BasicByColTest_ColIx1_Desc()
        {
            _sheet.Cells[1, 1].Value = 2;
            _sheet.Cells[2, 1].Value = 1;
            _sheet.Cells[3, 1].Value = 3;

            _sheet.Cells[1, 2].Value = "A";
            _sheet.Cells[2, 2].Value = "B";
            _sheet.Cells[3, 2].Value = "C";

            _sheet.Cells[4, 1].Formula = "SORT(A1:C3, 2, -1)";
            _sheet.Calculate();
            Assert.AreEqual(3, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(1, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[6, 1].Value);
            Assert.AreEqual("C", _sheet.Cells[4, 2].Value);
            Assert.AreEqual("B", _sheet.Cells[5, 2].Value);
            Assert.AreEqual("A", _sheet.Cells[6, 2].Value);
        }

        [TestMethod]
        public void BasicByRowTestAsc()
        {
            _sheet.Cells[1, 1].Value = 3;
            _sheet.Cells[2, 1].Value = 1;
            _sheet.Cells[3, 1].Value = 4;

            _sheet.Cells[1, 2].Value = 2;
            _sheet.Cells[2, 2].Value = 3;
            _sheet.Cells[3, 2].Value = 9;

            _sheet.Cells[1, 3].Value = 1;
            _sheet.Cells[2, 3].Value = 5;
            _sheet.Cells[3, 3].Value = 6;

            _sheet.Cells[4, 1].Formula = "SORT(A1:C3,1,1,FALSE)";
            _sheet.Calculate();

            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(3, _sheet.Cells[4, 2].Value);
            Assert.AreEqual(5, _sheet.Cells[4, 3].Value);
        }
    }
}
