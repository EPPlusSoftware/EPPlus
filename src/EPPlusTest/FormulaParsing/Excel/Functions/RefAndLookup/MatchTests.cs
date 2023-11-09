using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class MatchTests
    {
        private ParsingContext _parsingContext;
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void Match_Without_Wildcard()
        {
            _worksheet.Cells["A1"].Value = "test";
            _worksheet.Cells["A2"].Value = "value_to_match";
            _worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            _worksheet.Cells["A4"].Formula = "MATCH(\"value_to_match\", A1:A3, 0)";

            _worksheet.Calculate();

            Assert.AreEqual(2, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void Match_With_Wildcard1()
        {
            _worksheet.Cells["A1"].Value = "test";
            _worksheet.Cells["A2"].Value = "value_to_match";
            _worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            _worksheet.Cells["A4"].Formula = "MATCH(\"valu*_to_match\", A1:A3, 0)";

            _worksheet.Calculate();

            Assert.AreEqual(2, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void Match_With_Wildcard2()
        {
            _worksheet.Cells["A1"].Value = "test";
            _worksheet.Cells["A2"].Value = "value_to_match";
            _worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            _worksheet.Cells["A4"].Formula = "MATCH(\"?alue_to_match\", A1:A3, 0)";

            _worksheet.Calculate();

            Assert.AreEqual(2, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void Match_Without_ExactMatch()
        {
            _worksheet.Cells["A1"].Value = "test";
            _worksheet.Cells["A2"].Value = "value_to_match";
            _worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            _worksheet.Cells["A4"].Formula = "MATCH(\"no_match\", A1:A3, 0)";

            _worksheet.Calculate();

            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void Match_With_ExactMatch_ShouldNotTreatNullValuesAsAMatch()
        {
            _worksheet.Cells["A1"].Value = "test";
            _worksheet.Cells["A2"].Value = "value_to_match";
            _worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            _worksheet.Cells["A4"].Formula = "MATCH(B1, C1:C2, 0)";

            _worksheet.Calculate();

            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void MatchShouldReturnAnArray()
        {
            _worksheet.Cells["A1"].Value = "test";
            _worksheet.Cells["A2"].Value = "value_to_match";
            _worksheet.Cells["A3"].Value = "test";

            _worksheet.Cells["B1"].Value = "value_to_match";
            _worksheet.Cells["B2"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            _worksheet.Cells["A4"].Formula = "MATCH(B1:B2, A1:A3, 0)";

            _worksheet.Calculate();

            Assert.AreEqual(2, _worksheet.Cells["A4"].Value);
            Assert.AreEqual(1, _worksheet.Cells["A5"].Value);
        }
        [TestMethod]
        public void MatchShouldWorkWithSingleCell()
        {
            _worksheet.Cells["A1"].Value = "test@test.com";
            var formulaCell = _worksheet.Cells["B1"];
            formulaCell.Formula = "MATCH(\"?*@?*.?*\", A1, 0)";

            formulaCell.Calculate();

            Assert.AreEqual(1, formulaCell.GetValue<int>());
        }
    }
}
