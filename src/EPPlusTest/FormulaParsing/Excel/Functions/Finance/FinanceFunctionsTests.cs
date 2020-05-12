using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class FinanceFunctionsTests
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
        public void Npv_Tests()
        {
            _worksheet.Cells["A1"].Value = 0.02;
            _worksheet.Cells["A2"].Value = -5000;
            _worksheet.Cells["A3"].Value = 800;
            _worksheet.Cells["A4"].Value = 950;
            _worksheet.Cells["A5"].Value = 1080;
            _worksheet.Cells["A6"].Value = 1220;
            _worksheet.Cells["A7"].Value = 1500;

            _worksheet.Cells["A10"].Formula = "NPV(A1, A2:A7)";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A10"].Value, 2);
            Assert.AreEqual(196.88d, result);

            _worksheet.Cells["A1"].Value = 0.05;
            _worksheet.Cells["A2"].Value = -10000;
            _worksheet.Cells["A3"].Value = 2000;
            _worksheet.Cells["A4"].Value = 2400;
            _worksheet.Cells["A5"].Value = 2900;
            _worksheet.Cells["A6"].Value = 3500;
            _worksheet.Cells["A7"].Value = 4100;

            _worksheet.Cells["A10"].Formula = "NPV(A1, A3:A7) + A2";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A10"].Value, 2);
            Assert.AreEqual(2678.68, result);
        }

        [TestMethod]
        public void Fv_Tests()
        {
            _worksheet.Cells["A1"].Formula = "FV(5 %/ 12, 60, -1000)";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(68006.08d, result);

            _worksheet.Cells["A1"].Formula = "FV( 10%/4, 16, -2000, 0, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(39729.46, result);

            _worksheet.Cells["A1"].Formula = "FV(5%/12, 10 * 12, 0, -1000)";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 0);
            Assert.AreEqual(1647d, result);
        }

        [TestMethod]
        public void Pv_Tests()
        {
            _worksheet.Cells["A1"].Formula = "PV(5 %/ 12, 60, 1000)";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-52990.71, result);

            _worksheet.Cells["A1"].Formula = "PV( 10%/4, 16, 2000, 0, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-26762.76, result);
        }

        [TestMethod]
        public void Rate_Tests()
        {
            _worksheet.Cells["A1"].Formula = "RATE( 60, -1000, 50000 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(0.0062, result);

            _worksheet.Cells["A1"].Formula = "RATE( 24, -800, 0, 20000, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(0.0033, result);
        }
    }
}
