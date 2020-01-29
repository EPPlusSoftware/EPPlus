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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class InformationFunctionsTests
    {
        private ParsingContext _context;

        [TestInitialize]
        public void Setup()
        {
            _context = ParsingContext.Create();
        }

        [TestMethod]
        public void IsBlankShouldReturnTrueIfFirstArgIsNull()
        {
            var func = new IsBlank();
            var args = FunctionsHelper.CreateArgs(new object[]{null});
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsBlankShouldReturnTrueIfFirstArgIsEmptyString()
        {
            var func = new IsBlank();
            var args = FunctionsHelper.CreateArgs(string.Empty);
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsNumberShouldReturnTrueWhenArgIsNumeric()
        {
            var func = new IsNumber();
            var args = FunctionsHelper.CreateArgs(1d);
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsNumberShouldReturnfalseWhenArgIsNonNumeric()
        {
            var func = new IsNumber();
            var args = FunctionsHelper.CreateArgs("1");
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsErrorShouldReturnTrueIfArgIsAnErrorCode()
        {
            var args = FunctionsHelper.CreateArgs(ExcelErrorValue.Parse("#DIV/0!"));
            var func = new IsError();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsErrorShouldReturnFalseIfArgIsNotAnError()
        {
            var args = FunctionsHelper.CreateArgs("A", 1);
            var func = new IsError();
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsTextShouldReturnTrueWhenFirstArgIsAString()
        {
            var args = FunctionsHelper.CreateArgs("1");
            var func = new IsText();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsTextShouldReturnFalseWhenFirstArgIsNotAString()
        {
            var args = FunctionsHelper.CreateArgs(1);
            var func = new IsText();
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsNonTextShouldReturnFalseWhenFirstArgIsAString()
        {
            var args = FunctionsHelper.CreateArgs("1");
            var func = new IsNonText();
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsNonTextShouldReturnTrueWhenFirstArgIsNotAString()
        {
            var args = FunctionsHelper.CreateArgs(1);
            var func = new IsNonText();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsOddShouldReturnCorrectResult()
        {
            var args = FunctionsHelper.CreateArgs(3.123);
            var func = new IsOdd();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsEvenShouldReturnCorrectResult()
        {
            var args = FunctionsHelper.CreateArgs(4.123);
            var func = new IsEven();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsLogicalShouldReturnCorrectResult()
        {
            var func = new IsLogical();

            var args = FunctionsHelper.CreateArgs(1);
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);

            args = FunctionsHelper.CreateArgs("true");
            result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);

            args = FunctionsHelper.CreateArgs(false);
            result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void NshouldReturnCorrectResult()
        {
            var func = new N();

            var args = FunctionsHelper.CreateArgs(1.2);
            var result = func.Execute(args, _context);
            Assert.AreEqual(1.2, result.Result);

            args = FunctionsHelper.CreateArgs("abc");
            result = func.Execute(args, _context);
            Assert.AreEqual(0d, result.Result);

            args = FunctionsHelper.CreateArgs(true);
            result = func.Execute(args, _context);
            Assert.AreEqual(1d, result.Result);

            var errorCode = ExcelErrorValue.Create(eErrorType.Value);
            args = FunctionsHelper.CreateArgs(errorCode);
            result = func.Execute(args, _context);
            Assert.AreEqual(errorCode, result.Result);
        }
    }
}
