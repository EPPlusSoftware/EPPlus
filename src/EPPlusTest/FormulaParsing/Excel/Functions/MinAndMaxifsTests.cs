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
    public class MinAndMaxifsTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Init()
        {
            _package = new ExcelPackage();
            var sheet = _package.Workbook.Worksheets.Add("test");
            sheet.Cells["B3"].Value = "Hannah";
            sheet.Cells["C3"].Value = "F";
            sheet.Cells["D3"].Value = 93;
            sheet.Cells["B4"].Value = "Edward";
            sheet.Cells["C4"].Value = "M";
            sheet.Cells["D4"].Value = 79;
            sheet.Cells["B5"].Value = "Miranda";
            sheet.Cells["C5"].Value = "F";
            sheet.Cells["D5"].Value = 85;
            sheet.Cells["B6"].Value = "Miranda";
            sheet.Cells["C6"].Value = "F";
            sheet.Cells["D6"].Value = 82;
            sheet.Cells["B7"].Value = "William";
            sheet.Cells["C7"].Value = "M";
            sheet.Cells["D7"].Value = 64;
            _worksheet = sheet;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void MaxIfsShouldHandleOneCriteria()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"F\")";
            _worksheet.Calculate();
            Assert.AreEqual(93d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldHandleTwoCriterias()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"F\", B3:B7, \"Miranda\")";
            _worksheet.Calculate();
            Assert.AreEqual(85d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldHandleTwoCriteriasWithWildcards()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"F\", B3:B7, \"Mi**nda\")";
            _worksheet.Calculate();
            Assert.AreEqual(85d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldReturnZeroIfNoMatch()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"P\")";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldReturnValueErrorIfWrongSizeOnCriteriaRange()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"F\", B3:B5, \"Mi**nda\")";
            _worksheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value).ToString(), _worksheet.Cells["F1"].Value.ToString());
        }

        [TestMethod]
        public void MaxIfsShouldHandleNumericCriteriaWithOperator()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7, D3:D7,\">0\")";
            _worksheet.Calculate();
            Assert.AreEqual(93d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldHandleOneCriteria()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7,C3:C7,\"F\")";
            _worksheet.Calculate();
            Assert.AreEqual(82d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldHandleTwoCriterias()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7,C3:C7,\"F\", B3:B7, \"Miranda\")";
            _worksheet.Calculate();
            Assert.AreEqual(82d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldHandleTwoCriteriasWithWildcards()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7,C3:C7,\"F\", B3:B7, \"Mi**nda\")";
            _worksheet.Calculate();
            Assert.AreEqual(82d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldReturnZeroIfNoMatch()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7,C3:C7,\"P\")";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value); ;
        }

        [TestMethod]
        public void MinIfsShouldHandleNumericCriteriaWithOperator()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7, D3:D7,\">0\")";
            _worksheet.Calculate();
            Assert.AreEqual(64d, _worksheet.Cells["F1"].Value);
        }
        [TestMethod]
        public void MaxIfsShouldReturnZeroForCriteriaErrorNum()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7, D3:D7,#NUM!)";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldReturnZeroForCriteriaErrorValue()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7, D3:D7,#VALUE!)";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldReturnZeroForCriteriaWithBadOperand()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7, D3:D7,\">L\")";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldReturnZeroForCriteriaBlank()
        {
            _worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7, D3:D7, F2)";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }
        [TestMethod]
        public void MinIfsShouldReturnZeroForCriteriaErrorNum()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7, D3:D7,#NUM!)";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldReturnZeroForCriteriaErrorValue()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7, D3:D7,#VALUE!)";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldReturnZeroForCriteriaWithBadOperand()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7, D3:D7,\">L\")";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldReturnZeroForCriteriaBlank()
        {
            _worksheet.Cells["F1"].Formula = "MINIFS(D3:D7, D3:D7, F2)";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["F1"].Value);
        }
    }
}
