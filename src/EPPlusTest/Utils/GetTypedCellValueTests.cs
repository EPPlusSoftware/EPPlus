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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusTest.Utils
{
    [TestClass]
    public class GetTypedCellValueTests
    {
        [TestMethod]
        public void DoubleToNullableInt()
        {
            var result = ConvertUtil.GetTypedCellValue<int?>(2D);

            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void StringToDecimal()
        {
            var decimalSign = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            var result = ConvertUtil.GetTypedCellValue<decimal>($"1{decimalSign}4");

            Assert.AreEqual((decimal)1.4, result);
        }

        [TestMethod]
        public void EmptyStringToNullableDecimal()
        {
            var result = ConvertUtil.GetTypedCellValue<decimal?>("");
            Assert.IsNull(result);
        }

        [TestMethod]
        public void BlankStringToNullableDecimal()
        {
            var result = ConvertUtil.GetTypedCellValue<decimal?>(" ");

            Assert.IsNull(result);
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void EmptyStringToDecimal()
        {
            ConvertUtil.GetTypedCellValue<decimal>("");
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void FloatingPointStringToInt()
        {
            ConvertUtil.GetTypedCellValue<int>("1.4");
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void IntToDateTime()
        {
            ConvertUtil.GetTypedCellValue<DateTime>(122);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void IntToTimeSpan()
        {
            ConvertUtil.GetTypedCellValue<TimeSpan>(122);
        }

        [TestMethod]
        public void IntStringToTimeSpan()
        {
            Assert.AreEqual(TimeSpan.FromDays(122), ConvertUtil.GetTypedCellValue<TimeSpan>("122"));
        }

        [TestMethod]
        public void BoolToInt()
        {
            Assert.AreEqual(1, ConvertUtil.GetTypedCellValue<int>(true));
            Assert.AreEqual(0, ConvertUtil.GetTypedCellValue<int>(false));
        }

        [TestMethod]
        public void BoolToDecimal()
        {
            Assert.AreEqual(1m, ConvertUtil.GetTypedCellValue<decimal>(true));
            Assert.AreEqual(0m, ConvertUtil.GetTypedCellValue<decimal>(false));
        }

        [TestMethod]
        public void BoolToDouble()
        {
            Assert.AreEqual(1d, ConvertUtil.GetTypedCellValue<double>(true));
            Assert.AreEqual(0d, ConvertUtil.GetTypedCellValue<double>(false));
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void BadTextToInt()
        {
            ConvertUtil.GetTypedCellValue<int>("text1");
        }
        [TestMethod]
        public void DoubleToDateTime()
        {
            var expected = new DateTime(2020, 1, 1);
            Assert.AreEqual(expected, ConvertUtil.GetTypedCellValue<DateTime>(expected.ToOADate()));
        }
        [TestMethod]
        public void StringToDateTime()
        {
            var expected = new DateTime(2020, 1, 1);
            Assert.AreEqual(expected, ConvertUtil.GetTypedCellValue<DateTime>(expected.ToString()));
        }
        [TestMethod]
        public void DateTimeToTimeSpan()
        {
            var expected = new DateTime(2020, 1, 1);
            Assert.AreEqual(expected, ConvertUtil.GetTypedCellValue<DateTime>(new TimeSpan(expected.Ticks)));
        }
        [TestMethod]
        public void StringToTimeSpan()
        {
            var expected = new TimeSpan(10, 11, 12);
            Assert.AreEqual(expected, ConvertUtil.GetTypedCellValue<TimeSpan>(expected.ToString()));
        }
        [TestMethod]
        public void TimeSpanToDateTime()
        {
            var expected = new TimeSpan(10, 11, 12);
            Assert.AreEqual(expected, ConvertUtil.GetTypedCellValue<TimeSpan>(new DateTime(expected.Ticks)));
        }
        [TestMethod]
        public void DateTimeToNullableDateTime()
        {
            DateTime? expected = new DateTime(10, 11, 12);
            Assert.AreEqual(expected.Value, ConvertUtil.GetTypedCellValue<DateTime>(expected));
        }
        [TestMethod]
        public void DateTimeToNullableDateTimeNull()
        {
            DateTime? expected = null;
            Assert.AreEqual(default, ConvertUtil.GetTypedCellValue<DateTime>(expected));
        }
        [TestMethod]
        public void EmptyStringToNullableShouldReturnNull()
        {
            Assert.IsNull(ConvertUtil.GetTypedCellValue<int?>(""));
            Assert.IsNull(ConvertUtil.GetTypedCellValue<int?>("  "));

        }

        [TestMethod]
        public void TextToInt()
        {
            var result = ConvertUtil.GetTypedCellValue<int>("204");

            Assert.AreEqual(204, result);
        }
    }
}
