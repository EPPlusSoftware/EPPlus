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

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class FinanceFunctionTests
    {
        [TestMethod]
        public void PmtTest_EndOfPeriod()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "PMT( 5%/12, 60, 50000 )";
                sheet.Calculate();
                var value = sheet.Cells["A1"].Value;
                var value2 = System.Math.Round(Convert.ToDouble(value), 2);
                Assert.AreEqual(-943.56, value2);
            }
        }

        [TestMethod]
        public void PmtTest_BeginningOfPeriod()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "PMT( 5%/12, 60, 50000, 0, 1 )";
                sheet.Calculate();
                var value = sheet.Cells["A1"].Value;
                var value2 = System.Math.Round(Convert.ToDouble(value), 2);
                Assert.AreEqual(-939.65, value2);
            }
        }
    }
}
