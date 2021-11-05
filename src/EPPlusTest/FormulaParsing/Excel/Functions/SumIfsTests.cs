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
    public class SumIfsTests
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
            _sheet.Cells["A5"].Formula = "SUMIFS(A1:A4;B1:B5;\">5\";C1:C5;\">4\")";
            _sheet.Calculate();

            Assert.AreEqual(9d, _sheet.Cells["A5"].Value);
        }

        [TestMethod]
        public void ShouldIgnoreErrorInCriteriaRange()
        {
            _sheet.Cells["B3"].Value = ExcelErrorValue.Create(eErrorType.Div0);

            _sheet.Cells["A5"].Formula = "SUMIFS(A1:A4;B1:B5;\">5\";C1:C5;\">4\")";
            _sheet.Calculate();

            Assert.AreEqual(6d, _sheet.Cells["A5"].Value);
        }

        [TestMethod]
        public void ShouldHandleExcelRangesInCriteria()
        {
            _sheet.Cells["D1"].Value = 6;
            _sheet.Cells["A5"].Formula = "SUMIFS(A1:A4;B1:B5;\">5\";C1:C5;D1)";
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
    }
}
