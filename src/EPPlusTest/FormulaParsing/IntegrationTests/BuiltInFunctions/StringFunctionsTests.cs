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
    public class StringFunctionsTests : FormulaParserTestBase
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
        public void TextShouldConcatenateWithNextExpression()
        {
            var package = new ExcelPackage();
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetFormat(23.5, "$0.00")).Returns("$23.50");
            A.CallTo(() => provider.GetWorkbookNameValues()).Returns(new ExcelNamedRangeCollection(package.Workbook));
            var parser = new FormulaParser(provider);

            var result = parser.Parse("TEXT(23.5,\"$0.00\") & \" per hour\"");
            Assert.AreEqual("$23.50 per hour", result);
            package.Dispose();
        }

        [TestMethod]
        public void LenShouldAddLengthUsingSuppliedOperator()
        {
            var result = _parser.Parse("Len(\"abc\") + 2");
            Assert.AreEqual(5d, result);
        }

        [TestMethod]
        public void LowerShouldReturnALowerCaseString()
        {
            var result = _parser.Parse("Lower(\"ABC\")");
            Assert.AreEqual("abc", result);
        }

        [TestMethod]
        public void UpperShouldReturnAnUpperCaseString()
        {
            var result = _parser.Parse("Upper(\"abc\")");
            Assert.AreEqual("ABC", result);
        }

        [TestMethod]
        public void LeftShouldReturnSubstringFromLeft()
        {
            var result = _parser.Parse("Left(\"abacd\", 2)");
            Assert.AreEqual("ab", result);
        }

        [TestMethod]
        public void RightShouldReturnSubstringFromRight()
        {
            var result = _parser.Parse("RIGHT(\"abacd\", 2)");
            Assert.AreEqual("cd", result);
        }

        [TestMethod]
        public void MidShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Mid(\"abacd\", 2, 2)");
            Assert.AreEqual("ba", result);
        }

        [TestMethod]
        public void ReplaceShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Replace(\"testar\", 3, 3, \"hej\")");
            Assert.AreEqual("tehejr", result);
        }

        [TestMethod]
        public void SubstituteShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Substitute(\"testar testar\", \"es\", \"xx\")");
            Assert.AreEqual("txxtar txxtar", result);
        }

        [TestMethod]
        public void ConcatenateShouldReturnAccordingToParams()
        {
            var result = _parser.Parse("CONCATENATE(\"One\", \"Two\", \"Three\")");
            Assert.AreEqual("OneTwoThree", result);
        }

        [TestMethod]
        public void TShouldReturnText()
        {
            var result = _parser.Parse("T(\"One\")");
            Assert.AreEqual("One", result);
        }

        [TestMethod]
        public void ReptShouldConcatenate()
        {
            var result = _parser.Parse("REPT(\"*\",3)");
            Assert.AreEqual("***", result);
        }
    }
}
