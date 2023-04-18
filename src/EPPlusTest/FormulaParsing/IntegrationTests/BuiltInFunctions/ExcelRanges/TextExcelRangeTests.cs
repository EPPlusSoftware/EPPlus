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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges
{
    [TestClass]
    public class TextExcelRangeTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;
        private CultureInfo _currentCulture;

        [TestInitialize]
        public void Initialize()
        {
            _currentCulture = CultureInfo.CurrentCulture;
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("Test");

            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 3;
            _worksheet.Cells["A3"].Value = 6;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
            Thread.CurrentThread.CurrentCulture = _currentCulture;
        }

        [TestMethod]
        public void ExactShouldReturnTrueWhenEqualValues()
        {
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A4"].Formula = "EXACT(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void FindShouldReturnIndexCaseSensitive()
        {
            _worksheet.Cells["A1"].Value = "h";
            _worksheet.Cells["A2"].Value = "Hej hopp";
            _worksheet.Cells["A4"].Formula = "Find(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void FindShouldUse1basedIndex()
        {
            _worksheet.Cells["A4"].Formula = "Find(\"P\",\"P2\",1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void SearchShouldReturnIndexCaseInSensitive()
        {
            _worksheet.Cells["A1"].Value = "h";
            _worksheet.Cells["A2"].Value = "Hej hopp";
            _worksheet.Cells["A4"].Formula = "Search(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void SearchShouldUse1basedIndex()
        {
            _worksheet.Cells["A4"].Formula = "Search(\"P\",\"P2\",1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void ValueShouldHandleStringWithIntegers()
        {
            _worksheet.Cells["A1"].Value = "12";
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(12d, result);
        }

        [TestMethod]
        public void ValueShouldHandle1000delimiter()
        {
            var delimiter = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
            var val = $"5{delimiter}000";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(5000d, result);
        }

        [TestMethod]
        public void ValueShouldHandle1000DelimiterAndDecimal()
        {
            var delimiter = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            var val = $"5{delimiter}000{decimalSeparator}123";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(5000.123d, result);
        }

        [TestMethod]
        public void ValueShouldHandlePercent()
        {
            var val = $"20%";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(0.2d, result);
        }

        [TestMethod]
        public void ValueShouldHandleScientificNotation()
        {
            var func = new Value(new CultureInfo("en-US"));
            var arg = new FunctionArgument("1.2345E-02");
            var cr = func.Execute(new List<FunctionArgument> { arg }, ParsingContext.Create());
            Assert.AreEqual(0.012345d, cr.Result);
        }

        [TestMethod]
        public void ValueShouldHandleDate()
        {
            var ci = new CultureInfo("en-US");
            var func = new Value(ci);
            var date = new DateTime(2015, 12, 31);
            var arg = new FunctionArgument(date.ToString(ci));
            var cr = func.Execute(new List<FunctionArgument> { arg }, ParsingContext.Create());
            Assert.AreEqual(date.ToOADate(), cr.Result);
        }

        [TestMethod]
        public void ValueShouldHandleTime()
        {
            var ci = new CultureInfo("en-US");
            var func = new Value(ci);
            var date = new DateTime(2015, 12, 31);
            var date2 = new DateTime(2015, 12, 31, 12, 00, 00);
            var ts = date2.Subtract(date);
            var arg = new FunctionArgument(ts.ToString());
            var cr = func.Execute(new List<FunctionArgument> { arg }, ParsingContext.Create());
            Assert.AreEqual(0.5, cr.Result);
        }

        [TestMethod]
        public void ValueShouldReturn0IfValueIsNull()
        {

            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(0d, result);
        }

    }
}
