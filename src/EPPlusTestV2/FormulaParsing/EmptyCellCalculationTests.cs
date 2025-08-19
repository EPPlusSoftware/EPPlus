using EPPlusTest.FormulaParsing.IntegrationTests;
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class EmptyCellCalculationTests
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _package;
        private FormulaParser _parser;

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("Test");
            _parser = _package.Workbook.FormulaParser;

        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
            _parser = null;
        }

        [TestMethod]
        public void EmptyCellReferenceShouldReturnZero()
        {
            _sheet.Cells["A2"].Formula = "A1";
            _sheet.Calculate();
            var result = _sheet.Cells["A2"].Value;
            Assert.AreEqual(0d, result);
        }

        [TestMethod]
        public void EmptyCellReferenceMultiplicationShouldReturnZero()
        {
            _sheet.Cells["A2"].Formula = "A1*3";
            _sheet.Calculate();
            var result = _sheet.Cells["A2"].Value;
            Assert.AreEqual(0d, result);
        }

        [TestMethod]
        public void EmptyCellReferenceAdditionShouldReturnOtherOperand()
        {
            _sheet.Cells["A2"].Formula = "A1+2";
            _sheet.Calculate();
            var result = _sheet.Cells["A2"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void IfResultEmptyCellReferenceReturnsZero()
        {
            _sheet.Cells["A1"].Formula = "IF(TRUE,A2)";
            _sheet.Calculate();
            var result = _sheet.Cells["A1"].Value;
            Assert.AreEqual(0d, result);
        }

        [TestMethod]
        public void EmptyCellReferenceShouldEqualZero()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                sheet.Cells["A2"].Formula = "A1=0";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void EmptyCellReferenceShouldEqualEmptyString()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                sheet.Cells["A2"].Formula = "A1=\"\"";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void EmptyCellReferenceShouldEqualFalse()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                sheet.Cells["A2"].Formula = "A1=FALSE";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void IfConditionEmptyCellReferenceEqualsZero()
        {
            _sheet.Cells["A2"].Formula = "IF(A1=0,1)";
            _sheet.Calculate();
            var result = _sheet.Cells["A2"].Value;
            Assert.AreEqual(1d, result);
        }
        [TestMethod]
        public void IfConditionEmptyCellReferenceEqualsEmptyString()
        {
            _sheet.Cells["A2"].Formula = "IF(A1=\"\",1)";
            _sheet.Calculate();
            var result = _sheet.Cells["A2"].Value;
            Assert.AreEqual(1d, result);
        }
        [TestMethod]
        public void IfConditionEmptyCellReferenceEqualsFalse()
        {
            _sheet.Cells["A2"].Formula = "IF(A1=FALSE,1)";
            _sheet.Calculate();
            var result = _sheet.Cells["A2"].Value;
            Assert.AreEqual(1d, result);
        }
    }
}
