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
    public class UniqueTests
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
        public void ShouldReturnUniqueValuesByRow_1()
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[2, 1].Value = 1;
            _sheet.Cells[3, 1].Value = 3;

            _sheet.Cells[4, 1].Formula = "UNIQUE(A1:A3)";
            _sheet.Calculate();

            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(3, _sheet.Cells[5, 1].Value);
        }

        [TestMethod]
        public void ShouldReturnUniqueValuesByRow_2()
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[1, 2].Value = 2;
            _sheet.Cells[2, 1].Value = 1;
            _sheet.Cells[2, 2].Value = 2;
            _sheet.Cells[3, 1].Value = 3;
            _sheet.Cells[3, 2].Value = 4;

            _sheet.Cells[4, 1].Formula = "UNIQUE(A1:B3)";
            _sheet.Calculate();

            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[4, 2].Value);
            Assert.AreEqual(3, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(4, _sheet.Cells[5, 2].Value);
        }

        [TestMethod]
        public void ShouldReturnUniqueValuesByCol_1()
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[1, 2].Value = 1;
            _sheet.Cells[1, 3].Value = 3;

            _sheet.Cells[4, 1].Formula = "UNIQUE(A1:C1, TRUE)";
            _sheet.Calculate();

            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(3, _sheet.Cells[4, 2].Value);
        }

        [TestMethod]
        public void ShouldReturnUniqueValuesByCol_2()
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[2, 1].Value = 2;
            _sheet.Cells[1, 2].Value = 1;
            _sheet.Cells[2, 2].Value = 2;
            _sheet.Cells[1, 3].Value = 3;
            _sheet.Cells[2, 3].Value = 4;

            _sheet.Cells[4, 1].Formula = "UNIQUE(A1:C2,TRUE)";
            _sheet.Calculate();

            Assert.AreEqual(1, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(2, _sheet.Cells[5, 1].Value);
            Assert.AreEqual(3, _sheet.Cells[4, 2].Value);
            Assert.AreEqual(4, _sheet.Cells[5, 2].Value);
        }

        [TestMethod]
        public void UniqueShouldIgnoreCase()
        {
            _sheet.Cells[1, 1].Value = "A";
            _sheet.Cells[2, 1].Value = "a";
            _sheet.Cells[3, 1].Value = "B";

            _sheet.Cells[4, 1].Formula = "UNIQUE(A1:A3)";
            _sheet.Calculate();

            Assert.AreEqual("A", _sheet.Cells[4, 1].Value);
            Assert.AreEqual("B", _sheet.Cells[5, 1].Value);
        }

        [TestMethod]
        public void ShouldRemoveDuplicateColumns_WhenExactlyOnceIsSet()
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[2, 1].Value = 2;
            _sheet.Cells[1, 2].Value = 1;
            _sheet.Cells[2, 2].Value = 2;
            _sheet.Cells[1, 3].Value = 3;
            _sheet.Cells[2, 3].Value = 4;

            _sheet.Cells[4, 1].Formula = "UNIQUE(A1:C2,TRUE, TRUE)";
            _sheet.Calculate();

            Assert.AreEqual(3, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(4, _sheet.Cells[5, 1].Value);
            Assert.IsNull(_sheet.Cells[4, 2].Value);
            Assert.IsNull(_sheet.Cells[5, 2].Value);
        }

        [TestMethod]
        public void ShouldRemoveDuplicateRows_WhenExactlyOnceIsSet()
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[1, 2].Value = 2;
            _sheet.Cells[2, 1].Value = 1;
            _sheet.Cells[2, 2].Value = 2;
            _sheet.Cells[3, 1].Value = 3;
            _sheet.Cells[3, 2].Value = 4;

            _sheet.Cells[4, 1].Formula = "UNIQUE(A1:B3,FALSE,TRUE)";
            _sheet.Calculate();

            Assert.AreEqual(3, _sheet.Cells[4, 1].Value);
            Assert.AreEqual(4, _sheet.Cells[4, 2].Value);
            Assert.IsNull(_sheet.Cells[5, 1].Value);
            Assert.IsNull(_sheet.Cells[5, 2].Value);
        }
    }
}
