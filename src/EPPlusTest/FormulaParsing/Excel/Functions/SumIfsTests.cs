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
    }
}
