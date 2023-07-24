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
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml;
namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class SubtotalTests : FormulaParserTestBase
    {
        private ExcelWorksheet _worksheet;
        private ExcelPackage _package;

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage(new MemoryStream());
            _worksheet = _package.Workbook.Worksheets.Add("Test");
            _parser = _worksheet.Workbook.FormulaParser;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Avg()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(1, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Count()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(2, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_CountA()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(3, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Max()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(4, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Min()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(5, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Product()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(6, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Stdev()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(7, A2:A4)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(7, A5:A6)";
            _worksheet.Cells["A3"].Value = 5d;
            _worksheet.Cells["A4"].Value = 4d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Cells["A6"].Value = 4d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            result = Math.Round((double)result, 9);
            Assert.AreEqual(0.707106781d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_StdevP()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(8, A2:A4)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
            _worksheet.Cells["A3"].Value = 5d;
            _worksheet.Cells["A4"].Value = 4d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Cells["A6"].Value = 4d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(0.5d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Sum()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2m, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Var()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A4)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
            _worksheet.Cells["A3"].Value = 5d;
            _worksheet.Cells["A4"].Value = 4d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Cells["A6"].Value = 4d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(9m, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_VarP()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(10, A2:A4)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
            _worksheet.Cells["A3"].Value = 5d;
            _worksheet.Cells["A4"].Value = 4d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Cells["A6"].Value = 4d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(0.5d, result);
        }
    }
}
