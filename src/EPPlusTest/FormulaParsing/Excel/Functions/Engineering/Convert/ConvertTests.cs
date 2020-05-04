using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering.Convert
{
    [TestClass]
    public class ConvertTests
    {
        private ExcelPackage _package;
        private EpplusExcelDataProvider _provider;
        private ParsingContext _parsingContext;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _provider = new EpplusExcelDataProvider(_package);
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(RangeAddress.Empty);
            _worksheet = _package.Workbook.Worksheets.Add("testsheet");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void Bin2Dec_Tests()
        {
            _worksheet.Cells["A1"].Value = "10";
            _worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual(2, _worksheet.Cells["A2"].Value, "10 was not 2");

            _worksheet.Cells["A1"].Value = "11";
            _worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual(3, _worksheet.Cells["A2"].Value, "11 was not 3");

            _worksheet.Cells["A1"].Value = "0000000010";
            _worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual(2, _worksheet.Cells["A2"].Value, "0000000010 was not 2");

            _worksheet.Cells["A1"].Value = "1111111110";
            _worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual(-2, _worksheet.Cells["A2"].Value, "1111111110 was not -2");

            _worksheet.Cells["A1"].Value = 110;
            _worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual(6, _worksheet.Cells["A2"].Value, "110 was not 6");

            _worksheet.Cells["A1"].Value = 1110000110;
            _worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual(-122, _worksheet.Cells["A2"].Value, "110 was not 6");
        }

        [TestMethod]
        public void Bin2Hex_Tests()
        {
            _worksheet.Cells["A1"].Value = "10";
            _worksheet.Cells["A2"].Formula = "BIN2HEX(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("2", _worksheet.Cells["A2"].Value, "10 was not 2");

            _worksheet.Cells["A1"].Value = "0000000001";
            _worksheet.Cells["A2"].Formula = "BIN2HEX(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("1", _worksheet.Cells["A2"].Value, "0000000001 was not 1 but " + _worksheet.Cells["A2"].Value);

            _worksheet.Cells["A1"].Value = "1111111110";
            _worksheet.Cells["A2"].Formula = "BIN2HEX(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("FFFFFFFFFE", _worksheet.Cells["A2"].Value, "1111111110 was not FFFFFFFFFE but " + _worksheet.Cells["A2"].Value);

            _worksheet.Cells["A1"].Value = "11101";
            _worksheet.Cells["A2"].Formula = "BIN2HEX(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("1D", _worksheet.Cells["A2"].Value, "11101 was not 1D but " + _worksheet.Cells["A2"].Value);

            _worksheet.Cells["A1"].Value = "10";
            _worksheet.Cells["A2"].Formula = "BIN2HEX(A1,10)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("0000000002", _worksheet.Cells["A2"].Value, "10 (padded with 10) was not 0000000002 but " + _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void Bin2Oct_Tests()
        {
            _worksheet.Cells["A1"].Value = "101";
            _worksheet.Cells["A2"].Formula = "BIN2OCT(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("5", _worksheet.Cells["A2"].Value, "101 was not 5");

            _worksheet.Cells["A1"].Value = "0000000001";
            _worksheet.Cells["A2"].Formula = "BIN2OCT(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("1", _worksheet.Cells["A2"].Value, "0000000001 was not 1");

            _worksheet.Cells["A1"].Value = "10";
            _worksheet.Cells["A2"].Formula = "BIN2OCT(A1, 10)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("0000000002", _worksheet.Cells["A2"].Value, "10 was not 0000000002");

            _worksheet.Cells["A1"].Value = "1111111110";
            _worksheet.Cells["A2"].Formula = "BIN2OCT(A1, 10)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("7777777776", _worksheet.Cells["A2"].Value, "1111111110 was not 7777777776");

            _worksheet.Cells["A1"].Value = "1110";
            _worksheet.Cells["A2"].Formula = "BIN2OCT(A1)";
            _worksheet.Cells["A2"].Calculate();
            Assert.AreEqual("16", _worksheet.Cells["A2"].Value, "1110 was not 16");
        }
    }
}
