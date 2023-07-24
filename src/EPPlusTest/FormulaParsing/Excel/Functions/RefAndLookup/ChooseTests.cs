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
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class ChooseTests
    {
        private ParsingContext _parsingContext;
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _package = new ExcelPackage(new MemoryStream());
            _worksheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ChooseSingleValue()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "CHOOSE(4, A1, A2, A3, A4, A5)";
            _worksheet.Calculate();

            Assert.AreEqual(5d, _worksheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void ChooseSingleFormula()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "CHOOSE(6, A1, A2, A3, A4, A5, A6)";
            _worksheet.Calculate();

            Assert.AreEqual(12d, _worksheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void ChooseMultipleValues()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "SUM(CHOOSE({1,3,4}, A1, A2, A3, A4, A5))";
            _worksheet.Calculate();

            Assert.AreEqual(9M, _worksheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void ChooseValueAndFormula()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "SUM(CHOOSE({2,6}, A1, A2, A3, A4, A5, A6))";
            _worksheet.Calculate();

            Assert.AreEqual(14M, _worksheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void ChooseSumOfRange()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "SUM(CHOOSE(1, A1:A2, A2:A3))";
            _worksheet.Calculate();

            Assert.AreEqual(3M, _worksheet.Cells["B1"].Value);
        }

        private void fillChooseOptions()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 2d;
            _worksheet.Cells["A3"].Value = 3d;
            _worksheet.Cells["A4"].Value = 5d;
            _worksheet.Cells["A5"].Value = 7d;
            _worksheet.Cells["A6"].Formula = "A4 + A5";
        }
    }
}
