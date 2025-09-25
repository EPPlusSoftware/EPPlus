using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class MapTests : TestBase
    {
        [TestMethod]
        public void MapTest1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Value = 1;
            sheet.Cells["B2"].Value = 1;
            sheet.Cells["B3"].Value = 1;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["C2"].Value = 4;
            sheet.Cells["C3"].Value = 5;
            sheet.Cells["D1"].Value = 6;
            sheet.Cells["D2"].Value = 4;
            sheet.Cells["D3"].Value = 8;

            sheet.Cells["F5"].Formula = "MAP(A1:B3,C1:D3,LAMBDA(a,b,a+b))";
            sheet.Calculate();

            Assert.AreEqual(4d, sheet.Cells["F5"].Value);
            Assert.AreEqual(7d, sheet.Cells["G5"].Value);
            Assert.AreEqual(6d, sheet.Cells["F6"].Value);
            Assert.AreEqual(5d, sheet.Cells["G6"].Value);
            Assert.AreEqual(8d, sheet.Cells["F7"].Value);
            Assert.AreEqual(9d, sheet.Cells["G7"].Value);
        }

        [TestMethod]
        public void MapTest_ShouldHandleDifferentSizedRanges1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Value = 1;
            sheet.Cells["B2"].Value = 1;
            sheet.Cells["B3"].Value = 1;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["C2"].Value = 4;
            sheet.Cells["C3"].Value = 5;
            sheet.Cells["D1"].Value = 6;
            sheet.Cells["D2"].Value = 4;
            sheet.Cells["D3"].Value = 8;

            sheet.Cells["F5"].Formula = "MAP(A1:A3,C1:D3,LAMBDA(a,b,a+b))";
            sheet.Calculate();

            Assert.AreEqual(4d, sheet.Cells["F5"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G5"].Value);
            Assert.AreEqual(7d, sheet.Cells["F6"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G6"].Value);
            Assert.AreEqual(8d, sheet.Cells["F7"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G7"].Value);
        }

        [TestMethod]
        public void MapTest_ShouldHandleDifferentSizedRanges2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Value = 1;
            sheet.Cells["B2"].Value = 1;
            sheet.Cells["B3"].Value = 1;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["C2"].Value = 4;
            sheet.Cells["C3"].Value = 5;
            sheet.Cells["D1"].Value = 6;
            sheet.Cells["D2"].Value = 4;
            sheet.Cells["D3"].Value = 8;

            sheet.Cells["F5"].Formula = "MAP(A1:A2,C1:D3,LAMBDA(a,b,a+b))";
            sheet.Calculate();

            Assert.AreEqual(4d, sheet.Cells["F5"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G5"].Value);
            Assert.AreEqual(6d, sheet.Cells["F6"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G6"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["F7"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G7"].Value);

        }

        [TestMethod]
        public void MapTest_ShouldHandleDifferentSizedRanges3()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Value = 1;
            sheet.Cells["B2"].Value = 1;
            sheet.Cells["B3"].Value = 1;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["C2"].Value = 4;
            sheet.Cells["C3"].Value = 5;
            sheet.Cells["D1"].Value = 6;
            sheet.Cells["D2"].Value = 4;
            sheet.Cells["D3"].Value = 8;

            sheet.Cells["F5"].Formula = "MAP(A1:B2,C1:D3,LAMBDA(a,b,a+b))";
            sheet.Calculate();

            Assert.AreEqual(4d, sheet.Cells["F5"].Value);
            Assert.AreEqual(7d, sheet.Cells["G5"].Value);
            Assert.AreEqual(6d, sheet.Cells["F6"].Value);
            Assert.AreEqual(5d, sheet.Cells["G6"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["F7"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G7"].Value);
        }


        [TestMethod]
        public void Map_Xleta_Attribute_Abs()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = -1;
            sheet.Cells["A2"].Value = 0;
            sheet.Cells["A3"].Value = 5;

            sheet.Cells["B1"].Formula = "MAP(A1:A3, _xleta.ABS)";
            sheet.Calculate();

            Assert.AreEqual(1d, sheet.Cells["B1"].Value);   // ABS(-1)
            Assert.AreEqual(0d, sheet.Cells["B2"].Value);   // ABS(0)
            Assert.AreEqual(5d, sheet.Cells["B3"].Value);   // ABS(5)

        }

        [TestMethod]
        public void Map_Xleta_Attribute_Upper()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = "apple";
            sheet.Cells["A2"].Value = "banana";
            sheet.Cells["A3"].Value = "cherry";

            sheet.Cells["B1"].Formula = "MAP(A1:A3, _xleta.UPPER)";
            sheet.Calculate();

            Assert.AreEqual("APPLE", sheet.Cells["B1"].Value);
            Assert.AreEqual("BANANA", sheet.Cells["B2"].Value);
            Assert.AreEqual("CHERRY", sheet.Cells["B3"].Value);
        }

        [TestMethod]
        public void Map_Xleta_Attribute_Len()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = "dog";
            sheet.Cells["A2"].Value = "elephant";
            sheet.Cells["A3"].Value = "cat";

            sheet.Cells["B1"].Formula = "MAP(A1:A3, _xleta.LEN)";
            sheet.Calculate();

            Assert.AreEqual(3d, sheet.Cells["B1"].Value);   // "dog"
            Assert.AreEqual(8d, sheet.Cells["B2"].Value);   // "elephant"
            Assert.AreEqual(3d, sheet.Cells["B3"].Value);   // "cat"
        }

        [TestMethod]
        public void Map_Xleta_Attribute_IsNumber()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = 42;
            sheet.Cells["A2"].Value = "hello";
            sheet.Cells["A3"].Value = "";

            sheet.Cells["B1"].Formula = "MAP(A1:A3, _xleta.ISNUMBER)";
            sheet.Calculate();

            Assert.AreEqual(true, sheet.Cells["B1"].Value);   // 42 → TRUE
            Assert.AreEqual(false, sheet.Cells["B2"].Value);  // "hello" → FALSE
            Assert.AreEqual(false, sheet.Cells["B3"].Value);  // "" → FALSE
        }
    }
}
