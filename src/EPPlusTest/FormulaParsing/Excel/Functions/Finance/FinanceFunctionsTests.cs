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

        [TestMethod]
        public void Nper_Tests()
        {
            _worksheet.Cells["A1"].Formula = "NPER( 4%, -6000, 50000 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(10.34, result);

            _worksheet.Cells["A1"].Formula = "NPER( 6%/4, -2000, 60000, 30000, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(52.79, result);
        }

        [TestMethod]
        public void Irr_Tests()
        {
            _worksheet.Cells["B1"].Value = -100;
            _worksheet.Cells["B2"].Value = 20;
            _worksheet.Cells["B3"].Value = 24;
            _worksheet.Cells["B4"].Value = 28.80;
            _worksheet.Cells["B5"].Value = 34.56;
            _worksheet.Cells["B6"].Value = 41.47;

            _worksheet.Cells["C2"].Formula = "IRR(B1:B4)";
            _worksheet.Calculate();
            var result = System.Math.Round((double)_worksheet.Cells["C2"].Value, 2);
            Assert.AreEqual(-0.14, result);

            _worksheet.Cells["C2"].Formula = "IRR(B1:B6)";
            _worksheet.Calculate();
            result = System.Math.Round((double)_worksheet.Cells["C2"].Value, 2);
            Assert.AreEqual(0.13, result);
        }

        [TestMethod]
        public void Mirr_Tests()
        {
            _worksheet.Cells["B2"].Value = -100;
            _worksheet.Cells["B3"].Value = 18;
            _worksheet.Cells["B4"].Value = 22.5;
            _worksheet.Cells["B5"].Value = 28;
            _worksheet.Cells["B6"].Value = 35.5;
            _worksheet.Cells["B7"].Value = 45;

            _worksheet.Cells["C2"].Formula = "MIRR( B2:B6, 5.5%, 5% )";
            _worksheet.Calculate();
            var result = System.Math.Round((double)_worksheet.Cells["C2"].Value, 4);
            Assert.AreEqual(0.0254, result);

            _worksheet.Cells["C2"].Formula = "MIRR( B2:B7, 5.5%, 5% )";
            _worksheet.Calculate();
            result = System.Math.Round((double)_worksheet.Cells["C2"].Value, 1);
            Assert.AreEqual(0.1, result);
        }

        [TestMethod]
        public void Ipmt_Tests()
        {
            _worksheet.Cells["A1"].Formula = "IPMT( 5%/12, 1, 60, 50000 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-208.33, result);

            _worksheet.Cells["A1"].Formula = "IPMT( 5%/12, 2, 60, 50000 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-205.27, result);

            _worksheet.Cells["A1"].Formula = "IPMT( 3.5%/4, 1, 8, 0, 5000, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(0.00, result);

            _worksheet.Cells["A1"].Formula = "IPMT( 3.5%/4, 2, 8, 0, 5000, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(5.26, result);
        }

        [TestMethod]
        public void Ppmt_Tests()
        {
            _worksheet.Cells["A1"].Formula = "PPMT( 5%/12, 1, 60, 50000 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-735.23, result);

            _worksheet.Cells["A1"].Formula = "PPMT( 5%/12, 2, 60, 50000 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-738.29, result);

            _worksheet.Cells["A1"].Formula = "PPMT( 3.5%/4, 1, 8, 0, 5000, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-600.85, result);

            _worksheet.Cells["A1"].Formula = "PPMT( 3.5%/4, 2, 8, 0, 5000, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-606.11, result);
        }

        [TestMethod]
        public void Syd_Tests()
        {
            _worksheet.Cells["A1"].Formula = "SYD( 10000, 1000, 5, 1 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(3000d, result);

            _worksheet.Cells["A1"].Formula = "SYD( 10000, 1000, 5, 2 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(2400d, result);
        }

        [TestMethod]
        public void Sln_Tests()
        {
            _worksheet.Cells["A1"].Formula = "SLN( 10000, 1000, 5 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(1800d, result);

            _worksheet.Cells["A1"].Formula = "SLN( 500, 100, 8 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(50d, result);
        }
    }
}
