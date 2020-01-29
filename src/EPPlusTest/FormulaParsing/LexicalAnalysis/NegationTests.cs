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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;


namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class NegationTests
    {
        private SourceCodeTokenizer _tokenizer;

        [TestInitialize]
        public void Setup()
        {
            var context = ParsingContext.Create();
            _tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, null);
        }

        [TestCleanup]
        public void Cleanup()
        {

        }

        [TestMethod]
        public void ShouldSetNegatorOnFirstTokenIfFirstCharIsMinus()
        {
            var input = "-1";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(2, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.First().TokenType);
        }

        [TestMethod]
        public void ShouldChangePlusToMinusIfNegatorIsPresent()
        {
            var input = "1 + -1";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual(TokenType.Operator, tokens.ElementAt(1).TokenType);
            Assert.AreEqual("-", tokens.ElementAt(1).Value);
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInsideParenthethis()
        {
            var input = "1 + (-1 * 2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(8, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(3).TokenType);
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInsideFunctionCall()
        {
            var input = "Ceiling(-1, -0.1)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(8, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(2).TokenType);
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(5).TokenType, "Negator after comma was not identified");
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInEnumerable()
        {
            var input = "{-1}";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(1).TokenType);
        }

        [TestMethod]
        public void ShouldSetNegatorOnExcelAddress()
        {
            var input = "-A1";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(0).TokenType);
            Assert.AreEqual(TokenType.ExcelAddress, tokens.ElementAt(1).TokenType);
        }
    }
}
