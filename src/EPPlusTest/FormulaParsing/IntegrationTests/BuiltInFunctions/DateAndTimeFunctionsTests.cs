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
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using FakeItEasy;
using System.IO;
using System.Threading;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class DateAndTimeFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            _excelPackage = new ExcelPackage();
            _parser = new FormulaParser(_excelPackage);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _excelPackage.Dispose();
        }

        [TestMethod]
        public void DateShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Date(2012, 2, 2)");
            Assert.AreEqual(new DateTime(2012, 2, 2).ToOADate(), result);
        }

        [TestMethod]
        public void DateShouldHandleCellReference()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2012d;
                sheet.Cells["A2"].Formula = "Date(A1, 2, 2)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;
                Assert.AreEqual(new DateTime(2012, 2, 2).ToOADate(), result);
            }

        }

        [TestMethod]
        public void TodayShouldReturnAResult()
        {
            var result = _parser.Parse("Today()");
            Assert.IsInstanceOfType(DateTime.FromOADate((double)result), typeof(DateTime));
        }

        [TestMethod]
        public void NowShouldReturnAResult()
        {
            var result = _parser.Parse("now()");
            Assert.IsInstanceOfType(DateTime.FromOADate((double)result), typeof(DateTime));
        }

        [TestMethod]
        public void DayShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Day(Date(2012, 4, 2))");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void MonthShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Month(Date(2012, 4, 2))");
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void YearShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Year(Date(2012, 2, 2))");
            Assert.AreEqual(2012, result);
        }

        [TestMethod]
        public void TimeShouldReturnCorrectResult()
        {
            var expectedResult = ((double)(12 * 60 * 60 + 13 * 60 + 14))/((double)(24 * 60 * 60));
            var result = _parser.Parse("Time(12, 13, 14)");
            Assert.AreEqual(expectedResult, result);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResult()
        {
            var result = _parser.Parse("HOUR(Time(12, 13, 14))");
            Assert.AreEqual(12, result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResult()
        {
            var result = _parser.Parse("minute(Time(12, 13, 14))");
            Assert.AreEqual(13, result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Second(Time(12, 13, 59))");
            Assert.AreEqual(59, result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Second(\"10:12:14\")");
            Assert.AreEqual(14, result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Minute(\"10:12:14 AM\")");
            Assert.AreEqual(12, result);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Hour(\"10:12:14\")");
            Assert.AreEqual(10, result);
        }

        [TestMethod]
        public void DaysShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Days(Date(2015, 2, 2), Date(2015, 1, 1))");
            Assert.AreEqual(32d, result);
        }

        [TestMethod]
        public void Day360ShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Days360(Date(2012, 4, 2), Date(2012, 5, 2))");
            Assert.AreEqual(30, result);
        }

        [TestMethod]
        public void YearfracShouldReturnAResult()
        {
            var result = _parser.Parse("Yearfrac(Date(2012, 4, 2), Date(2012, 5, 2))");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void IsoWeekNumShouldReturnAResult()
        {
            var result = _parser.Parse("IsoWeekNum(Date(2012, 4, 2))");
            Assert.IsInstanceOfType(result, typeof(int));
        }

        [TestMethod]
        public void EomonthShouldReturnAResult()
        {
            var result = _parser.Parse("Eomonth(Date(2013, 2, 2), 3)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void WorkdayShouldReturnAResult()
        {
            var result = _parser.Parse("Workday(Date(2013, 2, 2), 3)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void DateNotEqualToStringShouldBeTrue()
        {
            var result = _parser.Parse("TODAY() <> \"\"");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void Calculation5()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "John";
            ws.Cells["B1"].Value = "Doe";
            ws.Cells["C1"].Formula = "B1&\", \"&A1";
            ws.Calculate();
            Assert.AreEqual("Doe, John", ws.Cells["C1"].Value);
        }

        [TestMethod]
        public void HourWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "HOUR(A1)";
            ws.Calculate();
            Assert.AreEqual(10, ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void MinuteWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "MINUTE(A1)";
            ws.Calculate();
            Assert.AreEqual(11, ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void SecondWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "SECOND(A1)";
            ws.Calculate();
            Assert.AreEqual(12, ws.Cells["B1"].Value);
        }
#if (!Core)
        [TestMethod]
        public void DateValueTest1()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "21 JAN 2015";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(2015, 1, 21).ToOADate(), ws.Cells["B1"].Value);
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void DateValueTestWithoutYear()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "21 JAN";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(currentYear, 1, 21).ToOADate(), ws.Cells["B1"].Value);
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void DateValueTestWithTwoDigitYear()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var expectedYear = 1960;
            ws.Cells["A1"].Value = "01/01/60";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(expectedYear, 1, 1).ToOADate(), ws.Cells["B1"].Value);
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void DateValueTestWithTwoDigitYear2()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var expectedYear = 2029;
            ws.Cells["A1"].Value = "01/01/29";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(expectedYear, 1, 1).ToOADate(), ws.Cells["B1"].Value);
            Thread.CurrentThread.CurrentCulture = ci;
        }


        [TestMethod]
        public void TimeValueTestPm()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "2:23 pm";
            ws.Cells["B1"].Formula = "TimeValue(A1)";
            ws.Calculate();
            var result = (double) ws.Cells["B1"].Value;
            Assert.AreEqual(0.599, Math.Round(result, 3));
            Thread.CurrentThread.CurrentCulture = ci;
        }


        [TestMethod]
        public void TimeValueTestFullDate()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "01/01/2011 02:23";
            ws.Cells["B1"].Formula = "TimeValue(A1)";
            ws.Calculate();
            var result = (double)ws.Cells["B1"].Value;
            Assert.AreEqual(0.099, Math.Round(result, 3));

            Thread.CurrentThread.CurrentCulture = ci;
        }
#endif
    }
}
