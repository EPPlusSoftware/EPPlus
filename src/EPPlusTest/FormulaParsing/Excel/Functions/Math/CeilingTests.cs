using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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
    public class CeilingTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(FormulaRangeAddress.Empty);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceLowerThan0()
        {
            var expectedValue = 22.36d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(22.35d, 0.01);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingBugTest1()
        {
            var expectedValue = 100d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(100d, 100d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }
        [TestMethod]
        public void CeilingBugTest2()
        {
            var expectedValue = 12000d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(12000d, 1000d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsMinus0point1()
        {
            var expectedValue = -22.4d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(-22.35d, -0.1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, System.Math.Round((double)result.Result, 2));
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs1()
        {
            var expectedValue = 23d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(22.35d, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs10()
        {
            var expectedValue = 30d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(22.35d, 10);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsNegative()
        {
            var expectedValue = -30d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(-22.35d, -10);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldThrowExceptionIfNumberIsPositiveAndSignificanceIsNegative()
        {
            var expectedValue = ExcelErrorValue.Parse("#NUM!");
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(22.35d, -1);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(expectedValue, result);
        }

        // CEILING.PRECISE
        [TestMethod]
        public void CeilingPreciseShouldHandleSingleArg()
        {
            var func = new CeilingPrecise();

            var args = FunctionsHelper.CreateArgs(6.1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(7d, result.Result);
        }

        [TestMethod]
        public void CeilingMathShouldReturnCorrectResult()
        {
            var func = new CeilingMath();

            var args = FunctionsHelper.CreateArgs(15.25);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(16d, result);

            args = FunctionsHelper.CreateArgs(15.25, 0.1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(15.3d, result);

            args = FunctionsHelper.CreateArgs(15.25, 5);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(20d, result);

            args = FunctionsHelper.CreateArgs(-15.25, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-15d, result);

            args = FunctionsHelper.CreateArgs(-15.25, 1, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-16d, result);

            args = FunctionsHelper.CreateArgs(-15.25, 10);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-10d, result);
        }
    }
}
