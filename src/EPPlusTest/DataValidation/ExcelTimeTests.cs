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
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ExcelTimeTests
    {
        private ExcelTime _time;
        private readonly decimal SecondsPerHour = 3600;
       // private readonly decimal HoursPerDay = 24;
        private readonly decimal SecondsPerDay = 3600 * 24;

        private decimal Round(decimal value)
        {
            return Math.Round(value, ExcelTime.NumberOfDecimals);
        }

        [TestInitialize]
        public void Setup()
        {
            _time = new ExcelTime();
        }

        [TestCleanup]
        public void Cleanup()
        {
            _time = null;
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void ExcelTimeTests_ConstructorWithValue_ShouldThrowIfValueIsLessThan0()
        {
            new ExcelTime(-1);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void ExcelTimeTests_ConstructorWithValue_ShouldThrowIfValueIsEqualToOrGreaterThan1()
        {
            new ExcelTime(1);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelTimeTests_Hour_ShouldThrowIfNegativeValue()
        {
            _time.Hour = -1;
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelTimeTests_Minute_ShouldThrowIfNegativeValue()
        {
            _time.Minute = -1;
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelTimeTests_Minute_ShouldThrowIValueIsGreaterThan59()
        {
            _time.Minute = 60;
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelTimeTests_Second_ShouldThrowIfNegativeValue()
        {
            _time.Second = -1;
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelTimeTests_Second_ShouldThrowIValueIsGreaterThan59()
        {
            _time.Second = 60;
        }

        [TestMethod]
        public void ExcelTimeTests_ToExcelTime_HourIsSet()
        {
            // Act
            _time.Hour = 1;
            
            // Assert
            Assert.AreEqual(Round(SecondsPerHour/SecondsPerDay), _time.ToExcelTime());
        }

        [TestMethod]
        public void ExcelTimeTests_ToExcelTime_MinuteIsSet()
        {
            // Arrange
            decimal expected = SecondsPerHour + (20M * 60M);
            // Act
            _time.Hour = 1;
            _time.Minute = 20;

            // Assert
            Assert.AreEqual(Round(expected/SecondsPerDay), _time.ToExcelTime());
        }

        [TestMethod]
        public void ExcelTimeTests_ToExcelTime_SecondIsSet()
        {
            // Arrange
            decimal expected = SecondsPerHour + (20M * 60M) + 10M;
            // Act
            _time.Hour = 1;
            _time.Minute = 20;
            _time.Second = 10;

            // Assert
            Assert.AreEqual(Round(expected / SecondsPerDay), _time.ToExcelTime());
        }

        [TestMethod]
        public void ExcelTimeTests_ConstructorWithValue_ShouldSetHour()
        {
            // Arrange
            decimal value = 3660M/(decimal)SecondsPerDay;

            // Act
            var time = new ExcelTime(value);

            // Assert
            Assert.AreEqual(1, time.Hour);
        }

        [TestMethod]
        public void ExcelTimeTests_ConstructorWithValue_ShouldSetMinute()
        {
            // Arrange
            decimal value = 3660M / (decimal)SecondsPerDay;

            // Act
            var time = new ExcelTime(value);

            // Assert
            Assert.AreEqual(1, time.Minute);
        }

        [TestMethod]
        public void ExcelTimeTests_ConstructorWithValue_ShouldSetSecond()
        {
            // Arrange
            decimal value = 3662M / (decimal)SecondsPerDay;

            // Act
            var time = new ExcelTime(value);

            // Assert
            Assert.AreEqual(2, time.Second);
        }

        [TestMethod]
        public void ExcelTimeTests_HourRoundingCheck()
        {
            decimal hour1 = decimal.Parse("0.416666666666667",CultureInfo.InvariantCulture);
            decimal hour2 = decimal.Parse("0.458333333333333",CultureInfo.InvariantCulture);

            var time1 = new ExcelTime(hour1);
            var time2 = new ExcelTime(hour2);

            Assert.AreEqual(10, time1.Hour);
            Assert.AreEqual(11, time2.Hour);
        }
    }
}
