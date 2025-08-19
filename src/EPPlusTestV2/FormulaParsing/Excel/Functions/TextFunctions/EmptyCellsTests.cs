using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class EmptyCellsTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;
        
        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ConcatenateShouldHandleEmptyCells()
        {
            _worksheet.Cells["A1"].Value = "A";
            _worksheet.Cells["C1"].Value = "C";
            _worksheet.Cells["A2"].Formula = "CONCATENATE(A1,B1,C1)";
            _worksheet.Calculate();
            Assert.AreEqual("AC", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void UpperShouldHandleEmptyCells()
        {
            _worksheet.Cells["A1"].Value = "A";
            _worksheet.Cells["C1"].Value = "C";
            _worksheet.Cells["A2"].Formula = "UPPER(B1)";
            _worksheet.Calculate();
            Assert.AreEqual("", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LowerShouldHandleEmptyCells()
        {
            _worksheet.Cells["A1"].Value = "A";
            _worksheet.Cells["C1"].Value = "C";
            _worksheet.Cells["A2"].Formula = "LOWER(B1)";
            _worksheet.Calculate();
            Assert.AreEqual("", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ProperHandleEmptyCells()
        {
            _worksheet.Cells["A1"].Value = "A";
            _worksheet.Cells["C1"].Value = "C";
            _worksheet.Cells["A2"].Formula = "PROPER(B1)";
            _worksheet.Calculate();
            Assert.AreEqual("", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LeftHandleEmptyCells()
        {
            _worksheet.Cells["A1"].Value = "A";
            _worksheet.Cells["C1"].Value = "C";
            _worksheet.Cells["A2"].Formula = "LEFT(B1,2)";
            _worksheet.Calculate();
            Assert.AreEqual("", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void RightHandleEmptyCells()
        {
            _worksheet.Cells["A1"].Value = "A";
            _worksheet.Cells["C1"].Value = "C";
            _worksheet.Cells["A2"].Formula = "RIGHT(B1,2)";
            _worksheet.Calculate();
            Assert.AreEqual("", _worksheet.Cells["A2"].Value);
        }
    }
}
