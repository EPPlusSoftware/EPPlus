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

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class LogicalFunctionsTests : FormulaParserTestBase
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
        public void IfShouldReturnCorrectResult()
        {
            var result = _parser.Parse("If(2 < 3, 1, 2)");
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void IIfShouldReturnCorrectResultWhenInnerFunctionExists()
        {
            var result = _parser.Parse("If(NOT(Or(true, FALSE)), 1, 2)");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void IIfShouldReturnCorrectResultWhenTrueConditionIsCoercedFromAString()
        {
            var result = _parser.Parse(@"If(""true"", 1, 2)");
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void IIfShouldReturnCorrectResultWhenFalseConditionIsCoercedFromAString()
        {
            var result = _parser.Parse(@"If(""false"", 1, 2)");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void NotShouldReturnCorrectResult()
        {
            var result = _parser.Parse("not(true)");
            Assert.IsFalse((bool)result);

            result = _parser.Parse("NOT(false)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void AndShouldReturnCorrectResult()
        {
            var result = _parser.Parse("And(true, 1)");
            Assert.IsTrue((bool)result);

            result = _parser.Parse("AND(true, true, 1, false)");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void OrShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Or(FALSE, 0)");
            Assert.IsFalse((bool)result);

            result = _parser.Parse("OR(true, true, 1, false)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void TrueShouldReturnCorrectResult()
        {
            var result = _parser.Parse("True()");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void FalseShouldReturnCorrectResult()
        {
            var result = _parser.Parse("False()");
            Assert.IsFalse((bool)result);
        }
    }
}
