using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class ReduceTests
    {
        [TestMethod]
        public void ReduceTest1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D5"].Formula = "REDUCE(,A1:C2,LAMBDA(a,b,a + b))";

            sheet.Calculate();

            Assert.AreEqual(21d, sheet.Cells["D5"].Value);
        }

        [TestMethod]
        public void ReduceTest_ShouldHandleArray()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["K2"].Value = 1;
            sheet.Cells["L2"].Value = 2;

            sheet.Cells["D5"].Formula = "REDUCE(K2:L2,A1:C2,LAMBDA(a,b,a + b))";

            sheet.Calculate();

            Assert.AreEqual(22d, sheet.Cells["D5"].Value);
            Assert.AreEqual(23d, sheet.Cells["E5"].Value);
        }

        [TestMethod]
        public void ReduceTest_Table()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "nums";
            sheet.Cells["A2"].Value = 77;
            sheet.Cells["A3"].Value = 2;
            sheet.Cells["A4"].Value = 3;
            sheet.Cells["A5"].Value = 4;
            sheet.Cells["A6"].Value = 5;
            sheet.Cells["A7"].Value = 56;
            sheet.Cells["A8"].Value = 6;
            sheet.Cells["A9"].Value = 7;
            sheet.Cells["A10"].Value = 78;
            sheet.Cells["A11"].Value = 8;

            sheet.Tables.Add(sheet.Cells["A1:A11"], "Table1");

            sheet.Cells["D5"].Formula = "REDUCE(1,Table1[nums],LAMBDA(a,b,IF(b>50,a*b,a)))";

            sheet.Calculate();

            Assert.AreEqual(336336d, sheet.Cells["D5"].Value);
        }

        [TestMethod]
        public void Reduce_Xleta_Attribute_Sum()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            // Mata in några värden i en kolumn
            sheet.Cells["A1"].Value = 5;
            sheet.Cells["A2"].Value = 10;
            sheet.Cells["A3"].Value = 15;

            // Använd REDUCE med eta-reducerad SUM
            sheet.Cells["C1"].Formula = "REDUCE(0, A1:A3, _xleta.SUM)";
            sheet.Calculate();

            // Förväntat resultat: 5 + 10 + 15 = 30
            Assert.AreEqual(30d, sheet.Cells["C1"].Value);
        }


        [TestMethod]
        public void Reduce_Xleta_Attribute_Average()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            // Mata in några värden i en kolumn
            sheet.Cells["A1"].Value = 5;
            sheet.Cells["A2"].Value = 10;
            sheet.Cells["A3"].Value = 15;

            // Använd REDUCE med eta-reducerad SUM
            sheet.Cells["C1"].Formula = "REDUCE(0, A1:A3, _xleta.AVERAGE)";
            sheet.Calculate();

            // Förväntat resultat: 5 + 10 + 15 = 30
            Assert.AreEqual(10.625d, sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void Reduce_Xleta_Attribute_Count_WithEmptyCells()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 5;
            sheet.Cells["A2"].Value = "";
            sheet.Cells["A3"].Value = 10;

            sheet.Cells["B1"].Formula = "REDUCE(1, A1:A3, _xleta.COUNT)";
            sheet.Calculate();

            // Expected: COUNT({1, 5}) = 2 → acc = 2
            // COUNT({2, ""}) = 1 → acc = 1
            // COUNT({1, 10}) = 2 → acc = 2
            Assert.AreEqual(2d, sheet.Cells["B1"].Value);
        }
    }
}
