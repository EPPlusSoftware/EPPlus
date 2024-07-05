﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class AreasTest
    {
        [TestMethod]
        public void AreashouldreturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A2"].Formula = "=AREAS(B2:D4)";
                sheet.Calculate();

                var result = sheet.Cells["A2"].Value;
                Assert.AreEqual(1, result);
            }
        }

        [TestMethod]
        public void AreashouldreturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B2"].Formula = "AREAS(A2:A3,A4:A5)";
                sheet.Calculate();

                var result = sheet.Cells["B2"].Value;
                Assert.AreEqual(2, result);
            }
        }
        [TestMethod]
        public void AreashouldreturnCorrectResult3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A2"].Formula = "=AREAS(B2:D4 B2 B2 B2 B2)";
                sheet.Calculate();

                var result = sheet.Cells["A2"].Value;
                Assert.AreEqual(1, result);
            }
        }

        [TestMethod]
        public void AreashouldreturnErrorNum()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A2"].Formula = "=AREAS(B2:D4 B2 B2 B2 B2 C1)";
                sheet.Calculate();

                var result = sheet.Cells["A2"].Value;
                Assert.AreEqual(ErrorValues.NullError, result);
            }
        }

        [TestMethod]
        public void AreashouldreturnCorrectResult4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B2"].Formula = "=AREAS((A1,A2,A3,A4,A5,A6,A7,A8,A9,A10:A12))";
                sheet.Calculate();

                var result = sheet.Cells["B2"].Value;
                Assert.AreEqual(10, result);
            }
        }
    }
}
