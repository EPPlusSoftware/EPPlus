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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestClass]
    public class CalcExtensionsTests
    {
        [TestMethod]
        public void ShouldCalculateChainTest()
        {
            var package = new ExcelPackage(new FileInfo("c:\\temp\\chaintest.xlsx"));
            package.Workbook.Calculate();
        }

        [TestMethod]
        public void CalculateTest()
        {
            //var pck = new ExcelPackage();
            //var ws = pck.Workbook.Worksheets.Add("Calc1");

            //ws.SetValue("A1", (short)1);
            //var v = pck.Workbook.FormulaParserManager.Parse("2.5-Calc1!A1+abs(3.0)-SIN(3)");
            //Assert.AreEqual(4.358879992, Math.Round((double)v, 9));

            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue("A1", (short)1);
            var v = pck.Workbook.FormulaParserManager.Parse("2.5-Calc1!A1+ABS(-3.0)-SIN(3)*abs(5)");
            Assert.AreEqual(3.79439996, Math.Round((double)v,9));
        }

        [TestMethod]
        public void CalculateTest2()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue("A1", (short)1);
            var v = pck.Workbook.FormulaParserManager.Parse("3*(2+5.5*2)+2*0.5+3");
            Assert.AreEqual(43, Math.Round((double)v, 9));
        }
    }
}
