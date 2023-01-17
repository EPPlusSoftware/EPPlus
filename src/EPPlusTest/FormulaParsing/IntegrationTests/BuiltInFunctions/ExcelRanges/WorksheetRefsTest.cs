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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges
{
    [TestClass]
    public class WorksheetRefsTest
    {
        private ExcelPackage _package;
        private ExcelWorksheet _firstSheet;
        private ExcelWorksheet _secondSheet;

        [TestInitialize]
        public void Init()
        {
            _package = new ExcelPackage();
            _firstSheet = _package.Workbook.Worksheets.Add("sheet1");
            _secondSheet = _package.Workbook.Worksheets.Add("sheet2");
            _firstSheet.Cells["A1"].Value = 1;
            _firstSheet.Cells["A2"].Value = 2;
        }

        [TestCleanup]
        public void Cleanup()
        {
            
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldHandleReferenceToOtherSheet()
        {
            _secondSheet.Cells["A1"].Formula = "SUM('sheet1'!A1:A2)";
            _secondSheet.Calculate();
            Assert.AreEqual(3d, _secondSheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldHandleReferenceToOtherSheetWithComplexName()
        {
            var sheet = _package.Workbook.Worksheets.Add("ab#k..2");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            _secondSheet.Cells["A1"].Formula = "SUM('ab#k..2'!A1:A2)";
            _secondSheet.Calculate();
            Assert.AreEqual(3d, _secondSheet.Cells["A1"].Value);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidFormulaException))]
        public void ShouldHandleInvalidRef()
        {
            var sheet = _package.Workbook.Worksheets.Add("ab#k..2");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            _secondSheet.Cells["A1"].Formula = "SUM('ab#k..2A1:A2')";
            _secondSheet.Calculate();
            Assert.IsInstanceOfType(_secondSheet.Cells["A1"].Value, typeof(ExcelErrorValue));
        }
    }
}
