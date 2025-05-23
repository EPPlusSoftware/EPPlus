using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class UnicodeTests
    {
        [TestMethod]
        public void LeftFunctionAndEqualsOperatorShouldHandleUnicode()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "🔼 test";
            ws.Cells[2, 1].Formula = "LEFT(A1, 1) = \"🔼\"";
            ws.Calculate(o => o.EnableUnicodeAwareStringOperations = true);
            Assert.IsInstanceOfType(ws.Cells[2, 1].Value, typeof(bool));
            Assert.IsTrue((bool)ws.Cells[2, 1].Value);
        }

        [TestMethod]
        public void RightFunctionAndEqualsOperatorShouldHandleUnicode()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "test 🔼";
            ws.Cells[2, 1].Formula = "RIGHT(A1, 1) = \"🔼\"";
            ws.Calculate(o => o.EnableUnicodeAwareStringOperations = true);
            Assert.IsInstanceOfType(ws.Cells[2, 1].Value, typeof(bool));
            Assert.IsTrue((bool)ws.Cells[2, 1].Value);
        }

        [TestMethod]
        public void EqualsOperatorShouldWorkWithUnicode()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Test 🔼";
            ws.Cells[2, 1].Formula = "A1=\"Test 🔼\"";
            ws.Calculate();
            Assert.IsInstanceOfType(ws.Cells[2, 1].Value, typeof(bool));
            Assert.IsTrue((bool)ws.Cells[2, 1].Value);
        }

        [TestMethod]
        public void NotEqualsOperatorShouldWorkWithUnicode()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Test 🔼";
            ws.Cells[2, 1].Formula = "A1<>\"Test 🔼\"";
            ws.Calculate();
            Assert.IsInstanceOfType(ws.Cells[2, 1].Value, typeof(bool));
            Assert.IsFalse((bool)ws.Cells[2, 1].Value);
        }

        [TestMethod]
        public void GreaterThanOperatorShouldWorkWithUnicode()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "🔼 test";
            ws.Cells[2, 1].Formula = "A1>\"S\"";
            ws.Calculate();
            Assert.IsInstanceOfType(ws.Cells[2, 1].Value, typeof(bool));
            Assert.IsTrue((bool)ws.Cells[2, 1].Value);
        }

        [TestMethod]
        public void LessThanOperatorShouldWorkWithUnicode()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "🔼 test";
            ws.Cells[2, 1].Formula = "A1<\"S\"";
            ws.Calculate();
            Assert.IsInstanceOfType(ws.Cells[2, 1].Value, typeof(bool));
            Assert.IsFalse((bool)ws.Cells[2, 1].Value);
        }
    }
}
