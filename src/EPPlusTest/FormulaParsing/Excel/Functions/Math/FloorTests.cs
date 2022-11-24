using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class FloorTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(FormulaRangeAddress.Empty);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIsBetween0And1()
        {
            var func = new Floor();
            var args = FunctionsHelper.CreateArgs(26.75d, 0.1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(26.7d, result.Result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIs1()
        {
            var func = new Floor();
            var args = FunctionsHelper.CreateArgs(26.75d, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(26d, result.Result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIsMinus1()
        {
            var func = new Floor();
            var args = FunctionsHelper.CreateArgs(-26.75d, -1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-26d, result.Result);
        }
        [TestMethod]
        public void FloorBugTest1()
        {
            var expectedValue = 100d;
            var func = new Floor();
            var args = FunctionsHelper.CreateArgs(100d, 100d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }
        [TestMethod]
        public void FloorBugTest2()
        {
            var expectedValue = 12000d;
            var func = new Floor();
            var args = FunctionsHelper.CreateArgs(12000d, 1000d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void FloorMathShouldReturnCorrectResult()
        {
            var func = new FloorMath();

            var args = FunctionsHelper.CreateArgs(58.55);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(58d, result);

            args = FunctionsHelper.CreateArgs(58.55, 0.1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(58.5d, result);

            args = FunctionsHelper.CreateArgs(58.55, 5);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(55d, result);

            args = FunctionsHelper.CreateArgs(-58.55, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-59d, result);

            args = FunctionsHelper.CreateArgs(-58.55, 1, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-58d, result);

            args = FunctionsHelper.CreateArgs(-58.55, 10);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-60d, result);
        }
    }
}
