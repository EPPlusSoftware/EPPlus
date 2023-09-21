using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class SequenceTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
        }

        [TestMethod]
        public void SequenceZeroRowArgument()
        {
            var expectedValue = ErrorValues.CalcError;
            var func = new Sequence();
            var args = FunctionsHelper.CreateArgs(0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void SequenceZeroColumnArgument()
        {
            var expectedValue = ErrorValues.CalcError;
            var func = new Sequence();
            var args = FunctionsHelper.CreateArgs(1, 0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void SequenceInvalidStep()
        {
            var func = new Sequence();
            var args = FunctionsHelper.CreateArgs(1, 1,1,"error string");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
        }

        [TestMethod]
        public void SequenceRow()
        {
            var func = new Sequence();
            var args = FunctionsHelper.CreateArgs(2,2, 5, 0.1);
            var result = func.Execute(args, _parsingContext);
            Assert.IsInstanceOfType(result.Result, typeof(InMemoryRange));
            var mr = (InMemoryRange)result.Result;
            Assert.AreEqual(2, mr.Size.NumberOfRows);
            Assert.AreEqual(2, mr.Size.NumberOfCols);
            Assert.AreEqual(5D, mr.GetValue(0, 0));
            Assert.AreEqual(5.1D, (double)mr.GetValue(0, 1), 0.0000001D);
            Assert.AreEqual(5.2D, (double)mr.GetValue(1, 0), 0.0000001D);
            Assert.AreEqual(5.3D, (double)mr.GetValue(1, 1), 0.0000001D);
        }
    }
}
