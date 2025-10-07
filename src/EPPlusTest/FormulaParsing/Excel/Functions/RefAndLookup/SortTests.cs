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

            _sheet.Cells[4, 1].Formula = "SORT(A1:B3, 2, 1, FALSE)";
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

            _sheet.Cells[4, 1].Formula = "SORT(A1:B3, 1, -1, FALSE)";
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

            _sheet.Cells[4, 1].Formula = "SORT(A1:C3, 2, -1, FALSE)";
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
            _sheet.Cells[2, 3].Value = 4;
            _sheet.Cells[3, 3].Value = 2;

            _sheet.Cells[4, 1].Formula = "SORT(A1:C3,1,1,TRUE)";
            _sheet.Calculate();

            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(4, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[6, 1].Value);

            Assert.AreEqual(2, _sheet.Cells[4, 2].Value);
            Assert.AreEqual(3, _sheet.Cells[5, 2].Value);
            Assert.AreEqual(9, _sheet.Cells[6, 2].Value);

            Assert.AreEqual(3, _sheet.Cells[4, 3].Value);
            Assert.AreEqual(1, _sheet.Cells[5, 3].Value);
            Assert.AreEqual(4, _sheet.Cells[6, 3].Value);
        }

        [TestMethod]
        public void BasicByRowTestAsc_Col1()
        {
            _sheet.Cells[1, 1].Value = 3;
            _sheet.Cells[2, 1].Value = 1;
            _sheet.Cells[3, 1].Value = 4;

            _sheet.Cells[1, 2].Value = 2;
            _sheet.Cells[2, 2].Value = 3;
            _sheet.Cells[3, 2].Value = 9;

            _sheet.Cells[1, 3].Value = 1;
            _sheet.Cells[2, 3].Value = 4;
            _sheet.Cells[3, 3].Value = 2;

            _sheet.Cells[4, 1].Formula = "SORT(A1:C3,3,1,TRUE)";
            _sheet.Calculate();

            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(4, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[6, 1].Value);

            Assert.AreEqual(3, _sheet.Cells[4, 2].Value);
            Assert.AreEqual(1, _sheet.Cells[5, 2].Value);
            Assert.AreEqual(4, _sheet.Cells[6, 2].Value);

            Assert.AreEqual(2, _sheet.Cells[4, 3].Value);
            Assert.AreEqual(3, _sheet.Cells[5, 3].Value);
            Assert.AreEqual(9, _sheet.Cells[6, 3].Value);
        }

        [TestMethod]
        public void NullValuesShouldAlwaysComeLast_1()
        {
            _sheet.Cells[1, 1].Value = 3;
            _sheet.Cells[2, 1].Value = 2;
            _sheet.Cells[3, 1].Value = 1;

            _sheet.Cells[1, 2].Value = "B";
            _sheet.Cells[2, 2].Value = "C";
            _sheet.Cells[3, 2].Value = "A";

            _sheet.Cells[5, 1].Formula = "SORT(A1:B4,1,-1)";
            _sheet.Calculate();

            Assert.AreEqual(3, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[6, 1].Value);
            Assert.AreEqual(1, _sheet.Cells[7, 1].Value);
            Assert.AreEqual(0D, _sheet.Cells[8, 1].Value);

        }

        [TestMethod]
        public void NullValuesShouldAlwaysComeLast_2()
        {
            _sheet.Cells[1, 1].Value = 3;
            _sheet.Cells[3, 1].Value = 2;
            _sheet.Cells[4, 1].Value = 1;

            _sheet.Cells[1, 2].Value = "B";
            _sheet.Cells[2, 2].Value = "C1";
            _sheet.Cells[3, 2].Value = "C";
            _sheet.Cells[4, 2].Value = "A";

            _sheet.Cells[5, 1].Formula = "SORT(A1:B4,1,-1)";
            _sheet.Calculate();

            Assert.AreEqual(3, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[6, 1].Value);
            Assert.AreEqual(1, _sheet.Cells[7, 1].Value);
            Assert.AreEqual(0D, _sheet.Cells[8, 1].Value);

        }

        [TestMethod]
        public void SortShouldHandleArrayWithIndexes()
        {
            _sheet.Cells[1, 1].Value = 150;
            _sheet.Cells[2, 1].Value = 150;
            _sheet.Cells[3, 1].Value = 150;
            _sheet.Cells[4, 1].Value = 150;
            _sheet.Cells[5, 1].Value = 150;
            _sheet.Cells[6, 1].Value = 150;

            _sheet.Cells[1, 2].Value = 2100;
            _sheet.Cells[2, 2].Value = 2110;
            _sheet.Cells[3, 2].Value = 2105;
            _sheet.Cells[4, 2].Value = 2100;
            _sheet.Cells[5, 2].Value = 2110;
            _sheet.Cells[6, 2].Value = 2105;

            _sheet.Cells[1, 3].Value = "Sub-Total for EUR";
            _sheet.Cells[2, 3].Value = "Total for G/L";
            _sheet.Cells[3, 3].Value = "Sub-Total for EUR";
            _sheet.Cells[4, 3].Value = "Total for G/L";
            _sheet.Cells[5, 3].Value = "Sub-Total for EUR";
            _sheet.Cells[6, 3].Value = "Total for G/L";

            _sheet.Cells[1, 3].Value = "EUR";
            _sheet.Cells[2, 3].Value = "[n/a]";
            _sheet.Cells[3, 3].Value = "EUR";
            _sheet.Cells[4, 3].Value = "[n/a]";
            _sheet.Cells[5, 3].Value = "EUR";
            _sheet.Cells[6, 3].Value = "[n/a]";

            _sheet.Cells["A10"].Formula = "SORT(A1:D6,{1,2},1,FALSE)";
            _sheet.Calculate();

            Assert.AreEqual(150, _sheet.Cells[10, 1].Value);
            Assert.AreEqual(150, _sheet.Cells[11, 1].Value);
            Assert.AreEqual(150, _sheet.Cells[12, 1].Value);
            Assert.AreEqual(150, _sheet.Cells[13, 1].Value);

            Assert.AreEqual(2100, _sheet.Cells[10, 2].Value, "Cell B10 was not 2100 as expected, result: " + _sheet.Cells[10, 2].Value);
            Assert.AreEqual(2100, _sheet.Cells[11, 2].Value, "Cell B11 was not 2100 as expected, result: " + _sheet.Cells[11, 2].Value);
            Assert.AreEqual(2105, _sheet.Cells[12, 2].Value, "Cell B12 was not 2105 as expected, result: " + _sheet.Cells[12, 2].Value);

        }

        [TestMethod]
        public void SortShouldHandleArrayWithIndexes2()
        {
            _sheet.Cells[1, 1].Value = 150;
            _sheet.Cells[2, 1].Value = 150;
            _sheet.Cells[3, 1].Value = 150;
            _sheet.Cells[4, 1].Value = 150;
            _sheet.Cells[5, 1].Value = 150;
            _sheet.Cells[6, 1].Value = 150;

            _sheet.Cells[1, 2].Value = 210000;
            _sheet.Cells[2, 2].Value = 210100;
            _sheet.Cells[3, 2].Value = 210050;
            _sheet.Cells[4, 2].Value = 210000;
            _sheet.Cells[5, 2].Value = 210100;
            _sheet.Cells[6, 2].Value = 210050;

            _sheet.Cells[1, 3].Value = "Sub-Total for EUR";
            _sheet.Cells[2, 3].Value = "Sub-Total for EUR";
            _sheet.Cells[3, 3].Value = "Sub-Total for EUR";
            _sheet.Cells[4, 3].Value = "Total for G/L";
            _sheet.Cells[5, 3].Value = "Total for G/L";
            _sheet.Cells[6, 3].Value = "Total for G/L";

            _sheet.Cells[1, 4].Value = "EUR";
            _sheet.Cells[2, 4].Value = "EUR";
            _sheet.Cells[3, 4].Value = "EUR";
            _sheet.Cells[4, 4].Value = "[n/a]";
            _sheet.Cells[5, 4].Value = "[n/a]";
            _sheet.Cells[6, 4].Value = "[n/a]";

            _sheet.Cells["A10"].Formula = "SORT(A1:D6,{1,2,3},1,FALSE)";
            _sheet.Calculate();

            Assert.AreEqual(150, _sheet.Cells[10, 1].Value);
            Assert.AreEqual(150, _sheet.Cells[11, 1].Value);
            Assert.AreEqual(150, _sheet.Cells[12, 1].Value);
            Assert.AreEqual(150, _sheet.Cells[13, 1].Value);

            Assert.AreEqual(210000, _sheet.Cells[10, 2].Value, "Cell B10 was not 2100 as expected, result: " + _sheet.Cells[10, 2].Value);
            Assert.AreEqual(210000, _sheet.Cells[11, 2].Value, "Cell B11 was not 2100 as expected, result: " + _sheet.Cells[11, 2].Value);
            Assert.AreEqual(210050, _sheet.Cells[12, 2].Value, "Cell B12 was not 2105 as expected, result: " + _sheet.Cells[12, 2].Value);
            Assert.AreEqual(210050, _sheet.Cells[13, 2].Value, "Cell B13 was not 2105 as expected, result: " + _sheet.Cells[13, 2].Value);

            Assert.AreEqual("Sub-Total for EUR", _sheet.Cells[10, 3].Value);
            Assert.AreEqual("Total for G/L", _sheet.Cells[11, 3].Value);
            Assert.AreEqual("Sub-Total for EUR", _sheet.Cells[12, 3].Value);
            Assert.AreEqual("Total for G/L", _sheet.Cells[13, 3].Value);

            Assert.AreEqual("EUR", _sheet.Cells[10, 4].Value);
            Assert.AreEqual("[n/a]", _sheet.Cells[11, 4].Value);

        }
    }
}
