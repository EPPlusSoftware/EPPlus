/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class SumIfsTests : TestBase
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;
        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            var s1 = _package.Workbook.Worksheets.Add("test");
            s1.Cells["A1"].Value = 1;
            s1.Cells["A2"].Value = 2;
            s1.Cells["A3"].Value = 3;
            s1.Cells["A4"].Value = 4;

            s1.Cells["B1"].Value = 5;
            s1.Cells["B2"].Value = 6;
            s1.Cells["B3"].Value = 7;
            s1.Cells["B4"].Value = 8;

            s1.Cells["C1"].Value = 5;
            s1.Cells["C2"].Value = 6;
            s1.Cells["C3"].Value = 7;
            s1.Cells["C4"].Value = 8;

            _sheet = s1;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldCalculateTwoCriteriaRanges()
        {
            _sheet.Cells["A5"].Formula = "SUMIFS(A1:A4,B1:B5,\">5\",C1:C5,\">4\")";
            _sheet.Calculate();

            Assert.AreEqual(9d, _sheet.Cells["A5"].Value);
        }

        [TestMethod]
        public void ShouldIgnoreErrorInCriteriaRange()
        {
            _sheet.Cells["B3"].Value = ExcelErrorValue.Create(eErrorType.Div0);

            _sheet.Cells["A5"].Formula = "SUMIFS(A1:A4,B1:B5,\">5\",C1:C5,\">4\")";
            _sheet.Calculate();

            Assert.AreEqual(6d, _sheet.Cells["A5"].Value);
        }

        [TestMethod]
        public void ShouldHandleExcelRangesInCriteria()
        {
            _sheet.Cells["D1"].Value = 6;
            _sheet.Cells["A5"].Formula = "SUMIFS(A1:A4,B1:B5,\">5\",C1:C5,D1)";
            _sheet.Calculate();

            Assert.AreEqual(2d, _sheet.Cells["A5"].Value);
        }

        [TestMethod]
        public void ShouldHandleTimeValuesCorrectly()
        {
            _sheet.Cells["A1"].Value = null;
            _sheet.Cells["A2"].Value = (7d * 3600d + 33d * 60d)/(24d * 3600d);// 07:33
            _sheet.Cells["A3"].Value = (11d * 3600d + 18d * 60d) / (24d * 3600d);// 11:18
            _sheet.Cells["A4"].Value = (7d * 3600d + 18d * 60d) / (24d * 3600d);// 07:18
            _sheet.Cells["A5"].Value = (10d * 3600d + 30d * 60d) / (24d * 3600d);// 10:30
            _sheet.Cells["A6"].Value = (10d * 3600d + 33d * 60d) / (24d * 3600d);// 10:33
            _sheet.Cells["A7"].Value = (10d * 3600d + 24d * 60d) / (24d * 3600d);// 10:24
            _sheet.Cells["A8"].Value = (11d * 3600d + 00d * 60d) / (24d * 3600d);// 11:00
            _sheet.Cells["A9"].Value = (6d * 3600d + 54d * 60d) / (24d * 3600d);// 06:54
            _sheet.Cells["A10"].Value = (12d * 3600d + 00d * 60d) / (24d * 3600d);// 12:00
            _sheet.Cells["A2:A10"].Calculate();

            for(var row = 2; row < 11; row++)
            {
                _sheet.Cells["B" + row].Value = 100;
            }

            _sheet.Cells["C2"].Formula = "SUMIFS(B:B,A:A,\">08:00\")";
            _sheet.Cells["C2"].Calculate();

            Assert.AreEqual(600d, _sheet.Cells["C2"].Value);

        }

        [TestMethod]
        public void SumIfsShouldHandleSingleRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "SUMIFS(H5,H5,\">0\",K5,\"> 0\")";
                sheet.Cells["H5"].Value = 1;
                sheet.Cells["K5"].Value = 1;
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A1"].Value);
            }
        }
        [TestMethod]
        public void ShouldHandleErrorInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[2, 1].Value = 4;
                sheet.Cells[3, 1].Value = 5;
                sheet.Cells[1, 2].Value = "#REF!";
                sheet.Cells[2, 2].Value = new ExcelErrorValue(eErrorType.Ref);
                sheet.Cells[3, 2].Value = 3;

                sheet.Cells[4, 1].Formula = "SUMIFS(A1:A3,B1:B3, #REF!)";
                sheet.Calculate();
                var val = sheet.Cells[4, 1].Value;
                Assert.AreEqual(4d, val);
            }
        }

        [TestMethod]
        public void SumIfsShouldIgnoreErrorsInRangeIfNotInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "a";
                sheet.Cells["B1"].Value = "b";
                sheet.Cells["C1"].Value = "c";
                sheet.Cells["A2"].Value = 1d;
                sheet.Cells["B2"].Value = ErrorValues.NAError;
                sheet.Cells["C2"].Value = "Test";

                sheet.Cells["A3"].Formula = "SUMIFS(A2:C2,A1:C1,\"=a\")";
                sheet.Calculate();

                Assert.AreEqual(1d, sheet.Cells["A3"].Value);
            }
        }
        [TestMethod]
        public void SumIfsShouldIgnoreErrorsInRangeIfInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["C1"].Value = 3;
                sheet.Cells["A2"].Value = "a";
                sheet.Cells["B2"].Value = ErrorValues.NAError;
                sheet.Cells["C2"].Value = "Test";

                sheet.Cells["A3"].Formula = "SUMIFS(A1:C1,A2:C2,\"=a\")";
                sheet.Calculate();

                Assert.AreEqual(1d, sheet.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void SumIfsShouldSumErrorsInRangeIfInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["C1"].Value = 3;
                sheet.Cells["A2"].Value = 1d;
                sheet.Cells["B2"].Value = ErrorValues.NAError;
                sheet.Cells["C2"].Value = "Test";

                sheet.Cells["A3"].Formula = "SUMIFS(A1:C1,A2:C2,\"=#n/a\")";
                sheet.Calculate();

                Assert.AreEqual(2d, sheet.Cells["A3"].Value);
            }
        }
        [TestMethod]
        public void SumIfsShouldSumIfCircularReferenceOusideOfInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2:A3"].Value = "Apples";
                sheet.Cells["A4:A5"].Value = "Artichokes";
                sheet.Cells["A6:A7"].Value = "Bananas";
                sheet.Cells["A8:A9"].Value = "Carrots";
                sheet.Cells["B2,B4,B6,B8"].Value = "Mats";
                sheet.Cells["B3,B5"].Value = "Jan";
                sheet.Cells["B7,B9"].Value = "Ossian";
                sheet.Cells["C2:C9"].FillNumber(10, 5);
                sheet.Cells["C5"].Formula = "=SUMIFS(C2:C9,A2:A9,\"=A*\",B2:B9,\"Mats\")";

                sheet.Calculate();
                Assert.AreEqual(sheet.Cells["C5"].Value, 10D + 20D);
            }
        }
        [TestMethod]
        public void SumIfsShouldSumIfCircularReferenceOutsideOfInCriteriaNotEquals()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2:A3"].Value = "Apples";
                sheet.Cells["A4:A5"].Value = "Artichokes";
                sheet.Cells["A6:A7"].Value = "Bananas";
                sheet.Cells["A8:A9"].Value = "Carrots";
                sheet.Cells["B2,B4,B6,B8"].Value = "Mats";
                sheet.Cells["B3,B5"].Value = "Jan";
                sheet.Cells["B7,B9"].Value = "Ossian";
                sheet.Cells["C2:C9"].FillNumber(10, 5);
                sheet.Cells["C5"].Formula = "=SUMIFS(C2:C9,A2:A9,\"<>A*\",B2:B9,\"Mats\")";

                sheet.Calculate();
                Assert.AreEqual(30D + 40D, sheet.Cells["C5"].Value);
            }
        }
        [TestMethod]
        public void SumIfsShouldHandleArraysInTheCriteriaRange_ColumnWise()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                LoadItemData(sheet);
                sheet.Cells["A2"].Value = "Crowbar";
                sheet.Cells["A3"].Value = "Hammer";
                sheet.Cells["A4"].Value = "Saw";
                sheet.Cells["B2"].Value = "Hammer";
                sheet.Cells["B3"].Value = "Butter";
                sheet.Cells["C2"].Formula = "=SUMIFS(N2:N11,K2:K11,A2:A11)";
                sheet.Cells["D2"].Formula = "=SUMIFS(N2:N11,K2:K11,A2:A3)";
                sheet.Cells["E2"].Formula = "=SUMIFS(N2:N11,K2:K11,A2:B3)";

                sheet.Calculate();

                Assert.AreEqual("C2:C11", sheet.Cells["C2"].FormulaRange.Address);
                Assert.AreEqual(270.6, (double)sheet.Cells["C2"].Value, 0.000001);
                Assert.AreEqual(88.2, sheet.Cells["C3"].Value);
                Assert.AreEqual(33.12, sheet.Cells["C4"].Value);
                Assert.AreEqual(0D, sheet.Cells["C5"].Value);
                Assert.AreEqual(0D, sheet.Cells["C11"].Value);
                Assert.IsNull(sheet.Cells["C12"].Value);

                Assert.AreEqual("D2:D3", sheet.Cells["D2"].FormulaRange.Address);
                Assert.AreEqual(270.6, (double)sheet.Cells["D2"].Value, 0.000001);
                Assert.AreEqual(88.2, sheet.Cells["D3"].Value);
                Assert.IsNull(sheet.Cells["D4"].Value);

                Assert.AreEqual("E2:F3", sheet.Cells["E2"].FormulaRange.Address);
                Assert.AreEqual(270.6, (double)sheet.Cells["E2"].Value, 0.000001);
                Assert.AreEqual(88.2, sheet.Cells["E3"].Value);
                Assert.IsNull(sheet.Cells["D4"].Value);
                Assert.AreEqual(88.2, (double)sheet.Cells["F2"].Value, 0.000001);
                Assert.AreEqual(7.2, sheet.Cells["F3"].Value);
                Assert.IsNull(sheet.Cells["F4"].Value);
            }
        }
        [TestMethod]
        public void SumIfsShouldHandleArraysInTheCriteriaRange_RowWise()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                LoadItemData(sheet);
                sheet.Cells["A2"].Value = "Crowbar";
                sheet.Cells["B2"].Value = "Hammer";
                sheet.Cells["C2"].Value = "Saw";
                sheet.Cells["A3"].Value = "Hammer";
                sheet.Cells["B3"].Value = "Butter";
                sheet.Cells["C5"].Formula = "=SUMIFS(N2:N11,K2:K11,A2:F2)";
                sheet.Cells["C6"].Formula = "=SUMIFS(N2:N11,K2:K11,A2:B2)";
                sheet.Cells["C7"].Formula = "=SUMIFS(N2:N11,K2:K11,A2:B3)";

                sheet.Calculate();

                Assert.AreEqual("C5:H5", sheet.Cells["C5"].FormulaRange.Address);
                Assert.AreEqual(270.6, (double)sheet.Cells["C5"].Value, 0.000001);
                Assert.AreEqual(88.2, sheet.Cells["D5"].Value);
                Assert.AreEqual(33.12, sheet.Cells["E5"].Value);
                Assert.AreEqual(0D, sheet.Cells["F5"].Value);
                Assert.AreEqual(0D, sheet.Cells["G5"].Value);
                Assert.IsNull(sheet.Cells["I5"].Value);

                Assert.AreEqual("C6:D6", sheet.Cells["C6"].FormulaRange.Address);
                Assert.AreEqual(270.6, (double)sheet.Cells["C6"].Value, 0.000001);
                Assert.AreEqual(88.2, sheet.Cells["D6"].Value);
                Assert.IsNull(sheet.Cells["E6"].Value);

                Assert.AreEqual("C7:D8", sheet.Cells["C7"].FormulaRange.Address);
                Assert.AreEqual(270.6, (double)sheet.Cells["C7"].Value, 0.000001);
                Assert.AreEqual(88.2, sheet.Cells["D7"].Value);
                Assert.IsNull(sheet.Cells["E7"].Value);
                Assert.AreEqual(88.2, (double)sheet.Cells["C8"].Value, 0.000001);
                Assert.AreEqual(7.2, sheet.Cells["D8"].Value);
                Assert.IsNull(sheet.Cells["E8"].Value);
            }
        }
        [TestMethod]
        public void SumIfsShouldHandleArraysWithMultipleCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                LoadItemData(sheet);
                sheet.Cells["A2"].Value = "Crowbar";
                sheet.Cells["A3"].Value = "Hammer";
                sheet.Cells["A4"].Value = "Saw";
                sheet.Cells["A5"].Value = "Monkey Wrench";
                sheet.Cells["B2"].Value = "Hardware";
                sheet.Cells["B3"].Value = "Software";
                sheet.Cells["B4"].Value = "Hardware";

                sheet.Cells["C2"].Formula = "SUMIFS(N2:N11,K2:K11,A2:A5,L2:L11,B2:B4)";
                sheet.Cells["D2"].Formula = "SUMIFS(N2:N11,K2:K11,A2:A5,N2:N11,\">50\")";

                sheet.Calculate();

                Assert.AreEqual("D2:D5", sheet.Cells["D2"].FormulaRange.Address);
                Assert.AreEqual(258.4, (double)sheet.Cells["D2"].Value, 0.000001);
                Assert.AreEqual(72.7D, sheet.Cells["D3"].Value);
                Assert.AreEqual(0D, sheet.Cells["D4"].Value);
                Assert.AreEqual(0D, sheet.Cells["D5"].Value);
                Assert.IsNull(sheet.Cells["D6"].Value);

                SaveWorkbook("SumIfsMultiArray.xlsx", package);
            }
        }

    }
}
