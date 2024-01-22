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
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class SourceCodeTokenizerTests
    {
        private SourceCodeTokenizer _tokenizer;

        [TestInitialize]
        public void Setup()
        {
            var context = ParsingContext.Create();
            _tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, OfficeOpenXml.FormulaParsing.NameValueProvider.Empty);
        }

        [TestCleanup]
        public void Cleanup()
        {
        }

        [TestMethod]
        public void ShouldCreateTokensForFunctionCorrectly()
        {
            var input = "Text(2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(4, tokens.Count());
            Assert.IsTrue(tokens.First().TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens.ElementAt(2).TokenTypeIsSet(TokenType.Integer));
            Assert.AreEqual("2", tokens.ElementAt(2).Value);
            Assert.IsTrue(tokens.Last().TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void ShouldHandleMultipleCharOperatorCorrectly()
        {
            var input = "1 <= 2";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual("<=", tokens.ElementAt(1).Value);
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.Operator));
        }

        [TestMethod]
        public void ShouldCreateTokensForEnumerableCorrectly()
        {
            var input = "Text({1;2})";
            var tokens = _tokenizer.Tokenize(input).ToArray();

            Assert.AreEqual(8, tokens.Length);
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.OpeningEnumerable));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.ClosingEnumerable));
        }

        [TestMethod]
        public void ShouldCreateTokensWithStringForEnumerableCorrectly()
        {
            var input = "{\"1\",\"2\"}";
            var tokens = _tokenizer.Tokenize(input).ToArray();

            Assert.AreEqual(5, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.OpeningEnumerable));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.StringContent));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.ClosingEnumerable));
        }

        [TestMethod]
        public void ShouldCreateTokensForExcelAddressCorrectly()
        {
            var input = "Text(A1)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.IsTrue(tokens.ElementAt(2).TokenTypeIsSet(TokenType.CellAddress));
        }

        [TestMethod]
        public void ShouldCreateTokenForPercentAfterDecimal()
        {
            var input = "1,23%";
            var tokens = _tokenizer.Tokenize(input);
            Assert.IsTrue(tokens.Last().TokenTypeIsSet(TokenType.Percent));
        }

        [TestMethod]
        public void TokenizerShouldIgnoreOperatorInString()
        {
            var input = "\"*\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.IsTrue(tokens.ElementAt(0).TokenTypeIsSet(TokenType.StringContent));
        }

        [TestMethod]
        public void TokenizerShouldHandleWorksheetNameWithMinus()
        {
            var input = "'A-B'!A1";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(5, tokens.Count());
            Assert.IsTrue(tokens.ElementAt(4).TokenTypeIsSet(TokenType.CellAddress));
        }

        [TestMethod]
        public void TestBug9_12_14()
        {
            //(( W60 -(- W63 )-( W29 + W30 + W31 ))/( W23 + W28 + W42 - W51 )* W4 )
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("test");
                for (var x = 1; x <= 10; x++)
                {
                    ws1.Cells[x, 1].Value = x;
                }

                ws1.Cells["A11"].Formula = "(( A1 -(- A2 )-( A3 + A4 + A5 ))/( A6 + A7 + A8 - A9 )* A5 )";
                //ws1.Cells["A11"].Formula = "(-A2 + 1 )";
                ws1.Calculate();
                var result = ws1.Cells["A11"].Value;
                Assert.AreEqual(-3.75, result);
            }
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSign()
        {
            var input = @"+3-3";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(3, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositive()
        {
            var input = @"-+3-3";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(4, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSignFromFirstFunctionArgument()
        {
            var input = @"SUM(+3-3,5)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Length);

            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSignFromSecondFunctionArgument()
        {
            var input = @"SUM(5,+3-3)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingDoubleNegatorFromFirstFunctionArgument()
        {
            var input = @"SUM(--3-3,5)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingDoubleNegatorFromSecondFunctionArgument()
        {
            var input = @"SUM(5,--3-3)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function), "TokenType was not function");
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis), "TokenType was not OpeningParenthesis");
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer), "TokenType was not Integer 2");
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Negator), "TokenType was not negator 5");
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesPositiveNegatorAsFirstFunctionArgument()
        {
            var input = @"SUM(+-3-3,5)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.AreEqual("-3", tokens[2].Value);
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositiveAsFirstFunctionArgument()
        {
            var input = @"SUM(-+3-3,5)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesPositiveNegatorAsSecondFunctionArgument()
        {
            var input = @"SUM(5,+-3-3)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.AreEqual("-3", tokens[4].Value);
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositiveAsSecondFunctionArgument()
        {
            var input = @"SUM(5,-+3-3)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }
        [TestMethod]
        public void TokenizeWorksheetName()
        {
            var input = @"sheetname!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(3, tokens.Length);
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.NameValue));
        }

        [TestMethod]
        public void TokenizeWorksheetNameWithQuotes()
        {
            var input = @"'sheetname'!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(5, tokens.Length);
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.WorksheetNameContent));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorksheetName()
        {
            var input = @"[0]sheetname!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(6, tokens.Length);
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.ExternalReference));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.WorksheetNameContent));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.NameValue));
        }

        [TestMethod]
        public void TokenizeExternalWorksheetNameWithQuotes()
        {
            var input = @"[3]'sheetname'!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Length);
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.ExternalReference));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.WorksheetNameContent));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorkbookName()
        {
            var input = @"[0]!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(5, tokens.Length);
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.ExternalReference));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorkbookInvalidRef()
        {
            var input = @"[0]#Ref!";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(4, tokens.Length);
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.InvalidReference));
        }
        [TestMethod]
        public void TokenizeExternalWorksheetInvalidRef()
        {
            var input = @"[0]Sheet1!#Ref!";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(6, tokens.Length);
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.InvalidReference));
        }

        [TestMethod]
        public void TokenizeShouldHandleWorksheetNameWithSingleQuote()
        {
            var input = @"=VLOOKUP(J7,'Sheet 1''21'!$Q$4:$R$28,2,0)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(17, tokens.Length);
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.WorksheetNameContent));
            Assert.AreEqual("Sheet 1''21", tokens[6].Value);
            
        }
        [TestMethod]
        public void NegatorHandlingWhenTokenizingIntegersAndAddresses()
        {
            var input = "10-'Sheet A'!A1";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(7,tokens.Length);
            Assert.AreEqual(TokenType.Operator, tokens[1].TokenType);
            Assert.AreEqual("-", tokens[1].Value);
        }
    }
}
