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
using ExGraph = OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionTree;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class FormulaParserTests
    {
        private FormulaParser _parser;

        [TestInitialize]
        public void Setup()
        {
            var provider = A.Fake<ExcelDataProvider>();
            _parser = new FormulaParser(provider);

        }

        [TestCleanup]
        public void Cleanup()
        {

        }

        [TestMethod]
        public void ParserShouldCallLexer()
        {
            var lexer = A.Fake<ILexer>();
            A.CallTo(() => (IEnumerable<Token>)lexer.Tokenize("ABC")).Returns(Enumerable.Empty<Token>());
            _parser.Configure(x => x.SetLexer(lexer));

            _parser.Parse("ABC");

            A.CallTo(() => lexer.Tokenize("ABC")).MustHaveHappened();
        }

        [TestMethod]
        public void ParserShouldCallGraphBuilder()
        {
            var lexer = A.Fake<ILexer>();
            var tokens = new List<Token>();
            A.CallTo(() => lexer.Tokenize("ABC")).Returns(tokens);
            var graphBuilder = A.Fake<IExpressionGraphBuilder>();
            A.CallTo(() => graphBuilder.Build(tokens)).Returns(new ExGraph());

            _parser.Configure(config =>
                {
                    config
                        .SetLexer(lexer)
                        .SetGraphBuilder(graphBuilder);
                });

            _parser.Parse("ABC");

            A.CallTo(() => graphBuilder.Build(tokens)).MustHaveHappened();
        }

        [TestMethod]
        public void ParserShouldCallCompiler()
        {
            var lexer = A.Fake<ILexer>();
            var tokens = new List<Token>();
            A.CallTo(() => lexer.Tokenize("ABC")).Returns(tokens);
            var ctx = ParsingContext.Create();
            var expectedGraph = new ExGraph();
            expectedGraph.Add(new StringExpression("asdf", ctx));
            var graphBuilder = A.Fake<IExpressionGraphBuilder>();
            A.CallTo(() => graphBuilder.Build(tokens)).Returns(expectedGraph);
            var compiler = A.Fake<IExpressionCompiler>();
            A.CallTo(() => compiler.Compile(expectedGraph.Expressions)).Returns(new CompileResult(0, DataType.Integer));

            _parser.Configure(config =>
            {
                config
                    .SetLexer(lexer)
                    .SetGraphBuilder(graphBuilder)
                    .SetExpresionCompiler(compiler);
            });

            _parser.Parse("ABC");

            A.CallTo(() => compiler.Compile(expectedGraph.Expressions)).MustHaveHappened();
        }

        [TestMethod]
        public void ParseAtShouldCallExcelDataProvider()
        {
            var excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => excelDataProvider.GetRangeFormula(string.Empty, 1, 1)).Returns("Sum(1,2)");
            var parser = new FormulaParser(excelDataProvider);
            var result = parser.ParseAt("A1");
            Assert.AreEqual(3d, result);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void ParseAtShouldThrowIfAddressIsNull()
        {
            _parser.ParseAt(null);
        }
    }
}
