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
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class TimeStringParserTests
    {
        private TimeStringParserV2 _parser;

        [TestInitialize]
        public void Initialize()
        {
            _parser = new TimeStringParserV2();
        }

        [TestCleanup]
        public void Cleanup() 
        {
            
            _parser = null;
        }
        private double GetSerialNumber(int hour, int minute, int second)
        {
            var secondsInADay = 24d * 60d * 60d;
            return ((double)hour * 60 * 60 + (double)minute * 60 + (double)second) / secondsInADay;
        }

        [TestMethod]
        public void ShouldParseTime_NoAmPm1()
        {
            var result = _parser.Parse("10:12:55");
            Assert.AreNotEqual(double.NaN, result, "Could not parse 10:12:55");
            var dt = DateTime.FromOADate(result);
            Assert.AreEqual(10, dt.Hour);
            Assert.AreEqual(12, dt.Minute);
            Assert.AreEqual(55, dt.Second);
        }

        [TestMethod]
        public void ShouldParseTime_NoAmPm2()
        {
            var result = _parser.Parse("22:12:55");
            Assert.AreNotEqual(double.NaN, result, "Could not parse 22:12:55");
            var dt = DateTime.FromOADate(result);
            Assert.AreEqual(22, dt.Hour);
            Assert.AreEqual(12, dt.Minute);
            Assert.AreEqual(55, dt.Second);
        }

        [TestMethod]
        public void ShouldParseTime_NoAmPm_HourAndMinuteOnly()
        {
            var result = _parser.Parse("13:12");
            Assert.AreNotEqual(double.NaN, result, "Could not parse 13:12");
            var dt = DateTime.FromOADate(result);
            Assert.AreEqual(13, dt.Hour);
            Assert.AreEqual(12, dt.Minute);
        }

        [TestMethod]
        public void CanParseShouldHandleValid12HourPatterns()
        {
            var parser = new TimeStringParserV2();
            Assert.AreNotEqual(double.NaN, parser.Parse("10:12:55 AM"), "Could not parse 10:12:55 AM");
            Assert.AreNotEqual(double.NaN, parser.Parse("9:12:55 PM"), "Could not parse 9:12:55 PM");
            Assert.AreNotEqual(double.NaN, parser.Parse("7 AM"), "Could not parse 7 AM");
            
        }

        [TestMethod]
        public void ShoulReturnNaN_WhenInvalidPmHour()
        {
            var result = _parser.Parse("14 AM");
            Assert.AreEqual(double.NaN, result);
        }

        [TestMethod]
        public void SerialNumberShouldBeCorrect()
        {
            var result = _parser.Parse("22:11:12.823");
            var roundedResult = Math.Round(result, 8);
            Assert.AreEqual(0.92445397, roundedResult);
        }

        [TestMethod]
        public void ShouldHandlePmCorrectly()
        {
            var result = _parser.Parse("4:12 PM");
            Assert.AreNotEqual(double.NaN, result, "Could not parse 4:12 PM");
            var dt = DateTime.FromOADate(result);
            Assert.AreEqual(16, dt.Hour);
            Assert.AreEqual(12, dt.Minute);
        }

        [TestMethod]
        public void ParseShouldIdentifyPatternAndReturnCorrectResult()
        {
            var parser = new TimeStringParserV2();
            var result = parser.Parse("10:12:55");
            Assert.AreEqual(GetSerialNumber(10, 12, 55), result);
        }

        [TestMethod]
        public void ParseShouldThrowExceptionIfSecondIsOutOfRange()
        {
            var parser = new TimeStringParserV2();
            var result = parser.Parse("10:12:60");
            Assert.AreEqual(result, double.NaN);
        }

        [TestMethod]
        public void ParseShouldThrowExceptionIfMinuteIsOutOfRange()
        {
            var parser = new TimeStringParserV2();
            var result = parser.Parse("10:60:55");
            Assert.AreEqual(result, double.NaN);
        }

        [TestMethod]
        public void ParseShouldIdentify12HourAMPatternAndReturnCorrectResult()
        {
            var parser = new TimeStringParserV2();
            var result = parser.Parse("10:12:55 AM");
            Assert.AreEqual(GetSerialNumber(10, 12, 55), result);
        }

        [TestMethod]
        public void ParseShouldIdentify12HourPMPatternAndReturnCorrectResult()
        {
            var parser = new TimeStringParserV2();
            var result = parser.Parse("10:12:55 PM");
            Assert.AreEqual(GetSerialNumber(22, 12, 55), result);
        }
    }
}
