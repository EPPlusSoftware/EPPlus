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
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestClass]
    public class BasicCalcTests : FormulaParserTestBase
    {
        private ExcelPackage _package;

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
            var excelDataProvider = new EpplusExcelDataProvider(_package);
            _parser = new FormulaParser(excelDataProvider);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldAddIntegersCorrectly()
        {
            var result = _parser.Parse("1 + 2");
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void ShouldSubtractIntegersCorrectly()
        {
            var result = _parser.Parse("2 - 1");
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void ShouldMultiplyIntegersCorrectly()
        {
            var result = _parser.Parse("2 * 3");
            Assert.AreEqual(6d, result);
        }

        [TestMethod]
        public void ShouldDivideIntegersCorrectly()
        {
            var result = _parser.Parse("8 / 4");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void ShouldDivideDecimalWithIntegerCorrectly()
        {
            var result = _parser.Parse("2.5/2");
            Assert.AreEqual(1.25d, result);
        }

        [TestMethod]
        public void ShouldHandleExpCorrectly()
        {
            var result = _parser.Parse("2 ^ 4");
            Assert.AreEqual(16d, result);
        }

        [TestMethod]
        public void ShouldHandleExpWithDecimalCorrectly()
        {
            var result = _parser.Parse("2.5 ^ 2");
            Assert.AreEqual(6.25d, result);
        }

        [TestMethod]
        public void ShouldMultiplyDecimalWithDecimalCorrectly()
        {
            var result = _parser.Parse("2.5 * 1.5");
            Assert.AreEqual(3.75d, result);
        }

        [TestMethod]
        public void ThreeGreaterThanTwoShouldBeTrue()
        {
            var result = _parser.Parse("3 > 2");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanTwoShouldBeFalse()
        {
            var result = _parser.Parse("3 < 2");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanOrEqualToThreeShouldBeTrue()
        {
            var result = _parser.Parse("3 <= 3");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanOrEqualToTwoDotThreeShouldBeFalse()
        {
            var result = _parser.Parse("3 <= 2.3");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void ThreeGreaterThanOrEqualToThreeShouldBeTrue()
        {
            var result = _parser.Parse("3 >= 3");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void TwoDotTwoGreaterThanOrEqualToThreeShouldBeFalse()
        {
            var result = _parser.Parse("2.2 >= 3");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void TwelveAndTwelveShouldBeEqual()
        {
            var result = _parser.Parse("2=2");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void TenPercentShouldBe0Point1()
        {
            var result = _parser.Parse("10%");
            Assert.AreEqual(0.1, result);
        }

        [TestMethod]
        public void ShouldHandleMultiplePercentSigns()
        {
            var result = _parser.Parse("10%%");
            Assert.AreEqual(0.001, result);
        }

        [TestMethod]
        public void ShouldHandlePercentageOnFunctionResult()
        {
            var result = _parser.Parse("SUM(1;2;3)%");
            Assert.AreEqual(0.06, result);
        }

        [TestMethod]
        public void ShouldHandlePercentageOnParantethis()
        {
            var result = _parser.Parse("(1+2)%");
            Assert.AreEqual(0.03, result);
        }

        [TestMethod]
        public void ShouldIgnoreLeadingPlus()
        {
            var result = _parser.Parse("+(1-2)");
            Assert.AreEqual(-1d, result);
        }

        [TestMethod]
        public void ShouldHandleDecimalNumberWhenDividingIntegers()
        {
            var result = _parser.Parse("224567455/400000000*500000");
            Assert.AreEqual(280709.31875, result);
        }

        [TestMethod]
        public void ShouldNegateExpressionInParenthesis()
        {
            var result = _parser.Parse("-(1+2)");
            Assert.AreEqual(-3d, result);
        }

        [TestMethod]
        public void ShouldHandlePercentStrings()
        {
            using(var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = "1%";
                sheet.Cells["B1"].Formula = "A1 * 5";
                sheet.Calculate();
                Assert.AreEqual(0.05d, sheet.Cells["B1"].Value);

                sheet.Cells["A1"].Value = "1%%";
                sheet.Cells["B1"].Formula = "A1 * 5";
                sheet.Calculate();
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
            }
        }
    }
}
