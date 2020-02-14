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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class ValueMatcherTests
    {
        private ValueMatcher _matcher;

        [TestInitialize]
        public void Setup()
        {
            _matcher = new ValueMatcher();
        }

        [TestMethod]
        public void ShouldReturn1WhenFirstParamIsSomethingAndSecondParamIsNull()
        {
            object o1 = 1;
            object o2 = null;
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void ShouldReturnMinus1WhenFirstParamIsNullAndSecondParamIsSomething()
        {
            object o1 = null;
            object o2 = 1;
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(-1, result);
        }

        [TestMethod]
        public void ShouldReturn0WhenBothParamsAreNull()
        {
            object o1 = null;
            object o2 = null;
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldReturn0WhenBothParamsAreEqual()
        {
            object o1 = 1d;
            object o2 = 1d;
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldReturnMinus1WhenFirstParamIsLessThanSecondParam()
        {
            object o1 = 1d;
            object o2 = 5d;
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(-1, result);
        }

        [TestMethod]
        public void ShouldReturn1WhenFirstParamIsGreaterThanSecondParam()
        {
            object o1 = 3d;
            object o2 = 1d;
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void ShouldReturn0WhenWhenParamsAreEqualStrings()
        {
            object o1 = "T";
            object o2 = "T";
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldReturn0WhenParamsAreEqualButDifferentTypes()
        {
            object o1 = "2";
            object o2 = 2d;
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(0, result, "IsMatch did not return 0 as expected when first param is a string and second a double");

            o1 = 2d;
            o2 = "2";
            result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(0, result, "IsMatch did not return 0 as expected when first param is a double and second a string");
        }

        [TestMethod]
        public void ShouldReturnIncompatibleOperandsWhenTypesDifferAndStringConversionToDoubleFails()
        {
            object o1 = 2d;
            object o2 = "T";
            var result = _matcher.IsMatch(o1, o2);
            Assert.AreEqual(ValueMatcher.IncompatibleOperands, result);
        }

        [TestMethod]
        public void ShouldReturn0WhenEqualDateTimeAndDouble()
        {
            var dt = new DateTime(2020, 2, 7).Date;
            var o2 = dt.ToOADate();
            var result = _matcher.IsMatch(dt, o2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldReturn1WhenDateTimeLargerThanDouble()
        {
            var dt = new DateTime(2020, 2, 7).Date;
            var o2 = dt.AddDays(-1).ToOADate();
            var result = _matcher.IsMatch(dt, o2);
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void ShouldReturn1WhenDateTimeSmallerThanDouble()
        {
            var dt = new DateTime(2020, 2, 7).Date;
            var o2 = dt.AddDays(1).ToOADate();
            var result = _matcher.IsMatch(dt, o2);
            Assert.AreEqual(-1, result);
        }
    }
}
