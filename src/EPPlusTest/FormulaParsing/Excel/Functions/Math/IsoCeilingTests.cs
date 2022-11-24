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
    public class IsoCeilingTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(FormulaRangeAddress.Empty);
        }

        [TestMethod]
        public void ShouldReturnCorrectResult()
        {
            var func = new IsoCeiling();

            var args = FunctionsHelper.CreateArgs(22.25);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(23d, result);

            args = FunctionsHelper.CreateArgs(22.25, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(23d, result);

            args = FunctionsHelper.CreateArgs(22.25, 0.1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(22.3d, result);

            args = FunctionsHelper.CreateArgs(22.25, 10);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(30d, result);

            args = FunctionsHelper.CreateArgs(-22.25, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-22d, result);

            args = FunctionsHelper.CreateArgs(-22.25, 0.1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-22.2d, result);

            args = FunctionsHelper.CreateArgs(-22.25, 5);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-20d, result);
        }
    }
}
