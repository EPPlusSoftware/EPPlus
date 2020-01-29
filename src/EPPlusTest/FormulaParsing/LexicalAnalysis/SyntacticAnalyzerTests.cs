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
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class SyntacticAnalyzerTests
    {
        private ISyntacticAnalyzer _analyser;

        [TestInitialize]
        public void Setup()
        {
            _analyser = new SyntacticAnalyzer();
        }

        [TestMethod]
        public void ShouldPassIfParenthesisAreWellformed()
        {
            var input = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("1", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis)
            };
            _analyser.Analyze(input);
        }

        [TestMethod, ExpectedException(typeof(FormatException))]
        public void ShouldThrowExceptionIfParenthesesAreNotWellformed()
        {
            var input = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("1", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            _analyser.Analyze(input);
        }

        [TestMethod]
        public void ShouldPassIfStringIsWellformed()
        {
            var input = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc123", TokenType.StringContent),
                new Token("'", TokenType.String)
            };
            _analyser.Analyze(input);
        }

        [TestMethod, ExpectedException(typeof(FormatException))]
        public void ShouldThrowExceptionIfStringHasNotClosing()
        {
            var input = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc123", TokenType.StringContent)
            };
            _analyser.Analyze(input);
        }


        [TestMethod, ExpectedException(typeof(UnrecognizedTokenException))]
        public void ShouldThrowExceptionIfThereIsAnUnrecognizedToken()
        {
            var input = new List<Token>
            {
                new Token("abc123", TokenType.Unrecognized)
            };
            _analyser.Analyze(input);
        }
    }
}
