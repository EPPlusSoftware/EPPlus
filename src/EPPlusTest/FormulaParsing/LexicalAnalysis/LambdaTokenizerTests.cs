using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class LambdaTokenizerTests
    {
        private ISourceCodeTokenizer _tokenizer;

        [TestInitialize]
        public void Setup()
        {
            //_tokenizer = SourceCodeTokenizer.Default;
            _tokenizer = SourceCodeTokenizer.Default;
        }


        [TestMethod]
        public void ShouldHandleInvokeArgs1()
        {
            var input = "LAMBDA(a, b, a + b)";
            var tokens = _tokenizer.Tokenize(input).ToArray();

            Assert.AreEqual(TokenType.Function, tokens[0].TokenType);
            Assert.AreEqual(TokenType.OpeningParenthesis, tokens[1].TokenType);
            Assert.AreEqual(TokenType.ParameterVariableDeclaration, tokens[2].TokenType);
            Assert.AreEqual(TokenType.Comma, tokens[3].TokenType);
            Assert.AreEqual(TokenType.ParameterVariableDeclaration, tokens[4].TokenType);
            Assert.AreEqual(TokenType.CommaLambda, tokens[5].TokenType);
            Assert.AreEqual(TokenType.ParameterVariable, tokens[6].TokenType);
            Assert.AreEqual(TokenType.Operator, tokens[7].TokenType);
            Assert.AreEqual(TokenType.ParameterVariable, tokens[8].TokenType);
            Assert.AreEqual(TokenType.ClosingParenthesis, tokens[9].TokenType);
        }

        [TestMethod]
        public void ShouldHandleInvokeArgs2()
        {
            var input = "IF(FALSE(),LAMBDA(r,r+2),A2:B5)(A2)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
        }
    }
}
