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
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Logging;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class TextFunctionsTests
    {
        [TestMethod]
        public void HyperlinkShouldHandleReference()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(B1)";
                sheet.Cells["B1"].Value = "http://epplus.codeplex.com";
                sheet.Calculate();
                Assert.AreEqual("http://epplus.codeplex.com", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void HyperlinkShouldHandleReference2()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(B1, B2)";
                sheet.Cells["B1"].Value = "http://epplus.codeplex.com";
                sheet.Cells["B2"].Value = "Epplus";
                sheet.Calculate();
                Assert.AreEqual("Epplus", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void HyperlinkShouldHandleText()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(\"testing\")";
                sheet.Calculate();
                Assert.AreEqual("testing", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void TrimShouldHandleStringWithSpaces()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "TRIM(B1)";
                sheet.Cells["B1"].Value = " epplus   5 ";
                sheet.Calculate();
                Assert.AreEqual("epplus 5", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void CharShouldReturnCharValOfNumber()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Char(A2)";
                sheet.Cells["A2"].Value = 55;
                sheet.Calculate();
                Assert.AreEqual("7", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void FixedShouldHaveCorrectDefaultValues()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2)";
                sheet.Cells["A2"].Value = 1234.5678;
                sheet.Calculate();
                Assert.AreEqual(1234.5678.ToString("N2"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void FixedShouldSetCorrectNumberOfDecimals()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,4)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.AreEqual(1234.56789.ToString("N4"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void FixedShouldSetNoCommas()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,4,true)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.AreEqual(1234.56789.ToString("F4"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void FixedShouldHandleNegativeDecimals()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,-1,true)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.AreEqual(1230.ToString("F0"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void ConcatenateShouldHandleRange()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Concatenate(1,A2)";
                sheet.Cells["A2"].Value = "hello";
                sheet.Calculate();
                Assert.AreEqual("1hello", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod, Ignore]
        public void Logtest1()
        {
            var sw = new Stopwatch();
            sw.Start();
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\denis.xlsx")))
            {
                var logger = LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\temp\log1.txt"));
                pck.Workbook.FormulaParser.Configure(x => x.AttachLogger(logger));
                pck.Workbook.Calculate();
                //
            }
            sw.Stop();
            var elapsed = sw.Elapsed;
            Console.WriteLine(string.Format("{0} seconds", elapsed.TotalSeconds));
        }
    }
}
