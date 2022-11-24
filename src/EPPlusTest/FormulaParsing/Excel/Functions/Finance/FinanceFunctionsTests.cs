using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
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
            _parsingContext.Scopes.NewScope(FormulaRangeAddress.Empty);
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
        public void Xirr_Tests1()
        {
            _worksheet.Cells["B2"].Value = -100;
            _worksheet.Cells["B3"].Value = 20;
            _worksheet.Cells["B4"].Value = 40;
            _worksheet.Cells["B5"].Value = 25;
            _worksheet.Cells["B6"].Value = 8;
            _worksheet.Cells["B7"].Value = 15;

            _worksheet.Cells["C2"].Value = new System.DateTime(2016, 01, 01);
            _worksheet.Cells["C3"].Value = new System.DateTime(2016, 04, 01);
            _worksheet.Cells["C4"].Value = new System.DateTime(2016, 10, 01);
            _worksheet.Cells["C5"].Value = new System.DateTime(2017, 02, 01);
            _worksheet.Cells["C6"].Value = new System.DateTime(2017, 03, 01);
            _worksheet.Cells["C7"].Value = new System.DateTime(2017, 06, 01);

            _worksheet.Cells["D2"].Formula = "XIRR(B2:B5, C2:C5)";
            _worksheet.Calculate();
            var result = System.Math.Round((double)_worksheet.Cells["D2"].Value, 4);
            Assert.AreEqual(-0.1967, result);

            _worksheet.Cells["D4"].Formula = "XIRR(B2:B7, C2:C7)";
            _worksheet.Calculate();
            result = System.Math.Round((double)_worksheet.Cells["D4"].Value, 4);
            Assert.AreEqual(0.0944, result);
        }

        [TestMethod]
        public void Xirr_Tests2()
        {
            _worksheet.Cells["A3"].Value = -10000;
            _worksheet.Cells["A4"].Value = 2750;
            _worksheet.Cells["A5"].Value = 4250;
            _worksheet.Cells["A6"].Value = 3250;
            _worksheet.Cells["A7"].Value = 2750;

            _worksheet.Cells["B3"].Value = new System.DateTime(2008, 01, 01);
            _worksheet.Cells["B4"].Value = new System.DateTime(2008, 03, 01);
            _worksheet.Cells["B5"].Value = new System.DateTime(2008, 10, 30);
            _worksheet.Cells["B6"].Value = new System.DateTime(2009, 02, 15);
            _worksheet.Cells["B7"].Value = new System.DateTime(2009, 04, 01);

            _worksheet.Cells["D2"].Formula = "XIRR(A3:A7, B3:B7, 0.1)";
            _worksheet.Calculate();
            var result = System.Math.Round((double)_worksheet.Cells["D2"].Value, 4);
            Assert.AreEqual(0.3734, result);
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
        public void Cumipmt_Tests()
        {
            _worksheet.Cells["A1"].Formula = "CUMIPMT( 5%/12, 60, 50000, 1, 12, 0 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-2294.98, result);

            _worksheet.Cells["A1"].Formula = "CUMIPMT( 5%/12, 60, 50000, 13, 24, 0 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-1833.10, result);

            _worksheet.Cells["A1"].Formula = "CUMIPMT( 5%/12, 60, 50000, 13, 24, 1 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-1825.49, result);
        }

        [TestMethod]
        public void Cumprinc_Tests()
        {
            _worksheet.Cells["A1"].Formula = "CUMPRINC( 5%/12, 60, 50000, 1, 12, 0  )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(-9027.7626, result);

            _worksheet.Cells["A1"].Formula = "CUMPRINC( 5%/12, 60, 50000, 13, 24, 0 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(-9489.6401, result);
        }

        [TestMethod]
        public void Ispmt_Tests()
        {
            _worksheet.Cells["A1"].Formula = "ISPMT( 5%/12, 1, 60, 50000 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-204.86, result);

            _worksheet.Cells["A1"].Formula = "ISPMT( 5%/12, 2, 60, 50000 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(-201.39, result);
        }

        [TestMethod]
        public void TestWorksheet()
        {
            using(var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("$Unit");
                var sheet = package.Workbook.Worksheets["$Unit"];
                Assert.IsNotNull(sheet);

                package.Workbook.Worksheets.Add("Unit1$");
                var sheet2 = package.Workbook.Worksheets["Unit1$"];
                Assert.IsNotNull(sheet2);
            }
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

        [TestMethod]
        public void Ddb_Tests()
        {
            _worksheet.Cells["A1"].Formula = "DDB( 10000, 1000, 5, 1 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(4000d, result);

            _worksheet.Cells["A1"].Formula = "DDB( 10000, 1000, 5, 4 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(864d, result);
        }

        [TestMethod]
        public void Db_Tests()
        {
            _worksheet.Cells["A1"].Formula = "DB( 10000, 1000, 5, 1, 6 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(1845d, result);

            _worksheet.Cells["A1"].Formula = "DB( 10000, 1000, 5, 5, 6 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(756.03d, result);

            _worksheet.Cells["A1"].Formula = "DB( 10000, 1000, 5, 6, 6 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(238.53d, result);
        }

        [TestMethod]
        public void FvSchedule_Tests()
        {
            _worksheet.Cells["B2"].Value = 0.05;
            _worksheet.Cells["B3"].Value = 0.05;
            _worksheet.Cells["B4"].Value = 0.035;
            _worksheet.Cells["B5"].Value = 0.035;
            _worksheet.Cells["B6"].Value = 0.035;
            _worksheet.Cells["A1"].Formula = "FVSCHEDULE( 10000, B2:B6 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(12223.61, result);

            _worksheet.Cells["A1"].Formula = "FVSCHEDULE( 1000, {0.02, 0.03, 0.04, 0.05} )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(1147.26, result);
        }

        [TestMethod]
        public void Pduration_Tests()
        {
            _worksheet.Cells["A1"].Formula = "PDURATION(4%, 10000, 15000)";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(10.34, result);

            _worksheet.Cells["A1"].Formula = "PDURATION(0.025/12,1000,1200)";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 1);
            Assert.AreEqual(87.6, result);
        }

        [TestMethod]
        public void Rri_Tests()
        {
            _worksheet.Cells["A1"].Formula = "RRI(10, 10000, 15000)";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(0.0414, result);
        }

        [TestMethod]
        public void Nominal_Tests()
        {
            _worksheet.Cells["A1"].Formula = "NOMINAL( 10%, 4 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(0.0965d, result);

            _worksheet.Cells["A1"].Formula = "NOMINAL( 2.5%, 12 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(0.0247d, result);
        }

        [TestMethod]
        public void Effect_Tests()
        {
            _worksheet.Cells["A1"].Formula = "EFFECT( 10%, 4 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(0.1038d, result);

            _worksheet.Cells["A1"].Formula = "EFFECT( 2.5%, 2 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(0.0252d, result);
        }

        [TestMethod]
        public void Xnpv_Tests()
        {
            _worksheet.Cells["B1"].Value = 0.05;

            _worksheet.Cells["A2"].Value = new DateTime(2016, 1, 1);
            _worksheet.Cells["A3"].Value = new DateTime(2016, 2, 1);
            _worksheet.Cells["A4"].Value = new DateTime(2016, 5, 1);
            _worksheet.Cells["A5"].Value = new DateTime(2016, 7, 1);
            _worksheet.Cells["A6"].Value = new DateTime(2016, 11, 1);
            _worksheet.Cells["A7"].Value = new DateTime(2017, 1, 1);

            _worksheet.Cells["B2"].Value = -10000;
            _worksheet.Cells["B3"].Value = 2000;
            _worksheet.Cells["B4"].Value = 2400;
            _worksheet.Cells["B5"].Value = 2900;
            _worksheet.Cells["B6"].Value = 3500;
            _worksheet.Cells["B7"].Value = 4100;

            _worksheet.Cells["A1"].Formula = "XNPV( B1, B2:B7, A2:A7 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(4447.94, result);
        }

        [TestMethod]
        public void Price_Tests()
        {
            _worksheet.Cells["B1"].Value = new DateTime(2012, 04, 01);
            _worksheet.Cells["B2"].Value = new DateTime(2020, 03, 31);
            _worksheet.Cells["A1"].Formula = "PRICE( B1, B2, 12%, 10%, 100, 2 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(110.8345, result);

            _worksheet.Cells["B1"].Value = new DateTime(2012, 04, 01);
            _worksheet.Cells["B2"].Value = new DateTime(2012, 06, 01);
            _worksheet.Cells["A1"].Formula = "PRICE( B1, B2, 12%, 10%, 100, 2 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(100.2623, result);
        }

        [TestMethod]
        public void YieldTest()
        {
            _worksheet.Cells["B1"].Value = new DateTime(2012, 01, 01);
            _worksheet.Cells["B2"].Value = new DateTime(2015, 06, 30);
            _worksheet.Cells["A1"].Formula = "YIELD( B1, B2, 10%, 101, 100, 4 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(0.0966, result);

            _worksheet.Cells["B1"].Value = new DateTime(2012, 01, 01);
            _worksheet.Cells["B2"].Value = new DateTime(2012, 01, 30);
            _worksheet.Cells["A1"].Formula = "YIELD( B1, B2, 10%, 101, 100, 4 )";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 4);
            Assert.AreEqual(-0.0235, result);
        }

        [TestMethod]
        public void YieldmatTest()
        {
            _worksheet.Cells["B1"].Value = new DateTime(2008, 03, 15);
            _worksheet.Cells["B2"].Value = new DateTime(2008, 11, 03);
            _worksheet.Cells["B3"].Value = new DateTime(2007, 11, 08);
            _worksheet.Cells["A1"].Formula = "YIELDMAT( B1, B2, B3, 6.25%, 100.0123, 0 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 6);
            Assert.AreEqual(0.060954, result);
        }

        [TestMethod]
        public void DurationTest()
        {
            _worksheet.Cells["B1"].Value = new DateTime(2015, 04, 01);
            _worksheet.Cells["B2"].Value = new DateTime(2025, 03, 31);
            _worksheet.Cells["A1"].Formula = "DURATION( B1, B2, 10%, 8%, 4 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(6.67, result);
        }

        [TestMethod]
        public void MdurationTest()
        {
            _worksheet.Cells["B1"].Value = new DateTime(2008, 01, 01);
            _worksheet.Cells["B2"].Value = new DateTime(2016, 01, 01);
            _worksheet.Cells["A1"].Formula = "MDURATION( B1, B2, 8%, 9%, 2 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 3);
            Assert.AreEqual(5.736, result);
        }

        [TestMethod]
        public void DiscTest()
        {
            _worksheet.Cells["B1"].Value = new DateTime(2016, 04, 01);
            _worksheet.Cells["B2"].Value = new DateTime(2021, 03, 31);
            _worksheet.Cells["A1"].Formula = "DISC( B1, B2, 95, 100 )";
            _worksheet.Calculate();

            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 2);
            Assert.AreEqual(0.01, result);

            _worksheet.Cells["B1"].Value = new DateTime(2018, 07, 01);
            _worksheet.Cells["B2"].Value = new DateTime(2048, 01, 01);
            _worksheet.Cells["A1"].Formula = "DISC(B1,B2,97.975,100,1)";
            _worksheet.Calculate();

            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 6);
            Assert.AreEqual(0.000686, result);
        }

        [TestMethod]
        public void DollarDeTest()
        {
            _worksheet.Cells["A1"].Formula = "DOLLARDE(1.99, 23)";
            _worksheet.Calculate();
            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 6);
            Assert.AreEqual(5.304348, result);

            _worksheet.Cells["A1"].Formula = "DOLLARDE(1.01, 16)";
            _worksheet.Calculate();
            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 6);
            Assert.AreEqual(1.0625, result);
        }

        [TestMethod]
        public void DollarFrTest()
        {
            _worksheet.Cells["A1"].Formula = "DOLLARFR(1.0625, 16)";
            _worksheet.Calculate();
            var result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 6);
            Assert.AreEqual(1.01, result);

            _worksheet.Cells["A1"].Formula = "DOLLARFR(1.09375, 32)";
            _worksheet.Calculate();
            result = System.Math.Round((double)_worksheet.Cells["A1"].Value, 6);
            Assert.AreEqual(1.03, result);
        }

        [TestMethod]
        public void IntrateTest1()
        {
            _worksheet.Cells["A1"].Value = new DateTime(2008, 2, 15);
            _worksheet.Cells["A2"].Value = new DateTime(2008, 5, 15);
            _worksheet.Cells["A3"].Formula = "INTRATE(A1, A2,1000000, 1014420, 2)";
            _worksheet.Calculate();
            var result = System.Math.Round((double)_worksheet.Cells["A3"].Value, 6);
            Assert.AreEqual(0.05768, result);
        }

        [TestMethod]
        public void IntrateTest2()
        {
            _worksheet.Cells["A1"].Value = new DateTime(2005, 4, 1);
            _worksheet.Cells["A2"].Value = new DateTime(2007, 3, 31);
            _worksheet.Cells["A3"].Formula = "INTRATE(A1, A2, 1000, 2125)";
            _worksheet.Calculate();
            var result = System.Math.Round((double)_worksheet.Cells["A3"].Value, 5);
            Assert.AreEqual(0.5625, result);
        }

        [TestMethod]
        public void TbilleqTest1()
        {
            _worksheet.Cells["A1"].Value = new DateTime(2008, 3, 31);
            _worksheet.Cells["A2"].Value = new DateTime(2008, 6, 1);
            _worksheet.Cells["A3"].Formula = "9.14%";
            _worksheet.Cells["A4"].Formula = "TBILLEQ(A1, A2, A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;

            Assert.AreEqual(0.09415149, System.Math.Round((double)result, 8));
        }

        [TestMethod]
        public void TbilleqTest2()
        {
            _worksheet.Cells["A1"].Value = new DateTime(2007, 8, 31);
            _worksheet.Cells["A2"].Value = new DateTime(2008, 6, 1);
            _worksheet.Cells["A3"].Formula = "9.14%";
            _worksheet.Cells["A4"].Formula = "TBILLEQ(A1, A2, A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;

            Assert.AreEqual(0.09800968, System.Math.Round((double)result, 8));
        }

        [TestMethod]
        public void TbillPriceTest1()
        {
            _worksheet.Cells["A1"].Value = new DateTime(2008, 3, 31);
            _worksheet.Cells["A2"].Value = new DateTime(2008, 6, 1);
            _worksheet.Cells["A3"].Formula = "9.14%";
            _worksheet.Cells["A4"].Formula = "TBILLPRICE(A1, A2, A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;

            Assert.AreEqual(98.42588889, System.Math.Round((double)result, 8));
        }

        [TestMethod]
        public void TbillPriceTest2()
        {
            _worksheet.Cells["A1"].Value = new DateTime(2007, 8, 31);
            _worksheet.Cells["A2"].Value = new DateTime(2008, 6, 1);
            _worksheet.Cells["A3"].Formula = "9.14%";
            _worksheet.Cells["A4"].Formula = "TBILLPRICE(A1, A2, A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;

            Assert.AreEqual(93.01805556, System.Math.Round((double)result, 8));
        }

        [TestMethod]
        public void TbillYieldTest1()
        {
            _worksheet.Cells["A1"].Value = new DateTime(2008, 3, 31);
            _worksheet.Cells["A2"].Value = new DateTime(2008, 6, 1);
            _worksheet.Cells["A3"].Value = 98.45;
            _worksheet.Cells["A4"].Formula = "TBILLYIELD(A1, A2, A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;

            Assert.AreEqual(0.09141696, System.Math.Round((double)result, 8));
        }

        [TestMethod]
        public void TbillYieldTest2()
        {
            _worksheet.Cells["A1"].Value = new DateTime(2007, 8, 31);
            _worksheet.Cells["A2"].Value = new DateTime(2008, 6, 1);
            _worksheet.Cells["A3"].Value = 98.45;
            _worksheet.Cells["A4"].Formula = "TBILLYIELD(A1, A2, A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;

            Assert.AreEqual(0.02061037, System.Math.Round((double)result, 8));
        }

        [DataTestMethod]
        [DataRow(2009, 1, 1, 2009, 12, 1, 2010, 4, 1, 75d, 4)]
        [DataRow(2009, 1, 1, 2009, 12, 1, 2010, 4, 18, 79.72222222d, 4)]
        [DataRow(2009, 1, 1, 2010, 4, 1, 2010, 4, 1, 25d, 4)]
        [DataRow(2009, 1, 1, 2010, 7, 1, 2010, 4, 1, 0d, 4)]
        [DataRow(2009, 1, 1, 2010, 7, 1, 2010, 4, 5, -1.11111111d, 4)]
        [DataRow(2009, 1, 1, 2010, 10, 1, 2010, 4, 1, -25d, 4)]
        [DataRow(2009, 1, 1, 2009, 12, 1, 2010, 4, 1, 125d, 2)]
        [DataRow(2009, 1, 1, 2009, 12, 1, 2010, 4, 1, 125d, 1)]
        public void AccrintTestCalcFirstInterest(int iYear, int iMonth, int iDay, int fiYear, int fiMonth, int fiDay, int sYear, int sMonth, int sDay, double expectedResult, int frequency)
        {
            var issue = new DateTime(iYear, iMonth, iDay);
            var firstInterest = new DateTime(fiYear, fiMonth, fiDay);
            var settlement = new DateTime(sYear, sMonth, sDay);
            _worksheet.Cells["A1"].Value = issue;
            _worksheet.Cells["A2"].Value = firstInterest;
            _worksheet.Cells["A3"].Value = settlement;
            _worksheet.Cells["A4"].Value = 0.1;
            _worksheet.Cells["A5"].Value = 1000;
            _worksheet.Cells["A6"].Value = frequency;
            _worksheet.Cells["A7"].Value = 0;
            _worksheet.Cells["A9"].Formula = "ACCRINT(A1, A2, A3, A4, A5, A6, A7, 0)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A9"].Value;

            Assert.AreEqual(expectedResult, System.Math.Round((double)result, 8));
        }

        [DataTestMethod]
        [DataRow(2009, 1, 1, 2009, 12, 1, 2010, 4, 1, 125d, 4)]
        [DataRow(2009, 2, 15, 2009, 12, 1, 2010, 4, 1, 112.77777778d, 4)]
        [DataRow(2009, 2, 15, 2009, 12, 1, 2009, 4, 2, 13.05555556d, 4)]
        public void AccrintIssueToSettlementTest(int iYear, int iMonth, int iDay, int fiYear, int fiMonth, int fiDay, int sYear, int sMonth, int sDay, double expectedResult, int frequency)
        {
            var issue = new DateTime(iYear, iMonth, iDay);
            var firstInterest = new DateTime(fiYear, fiMonth, fiDay);
            var settlement = new DateTime(sYear, sMonth, sDay);
            _worksheet.Cells["A1"].Value = issue;
            _worksheet.Cells["A2"].Value = firstInterest;
            _worksheet.Cells["A3"].Value = settlement;
            _worksheet.Cells["A4"].Value = 0.1;
            _worksheet.Cells["A5"].Value = 1000;
            _worksheet.Cells["A6"].Value = 2;
            _worksheet.Cells["A7"].Value = frequency;
            _worksheet.Cells["A9"].Formula = "ACCRINT(A1, A2, A3, A4, A5, A6, A7)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A9"].Value;

            Assert.AreEqual(expectedResult, System.Math.Round((double)result, 8));
        }

        [DataTestMethod]
        [DataRow(2009, 1, 1, 2009, 4, 1, 25d, 0)]
        [DataRow(2009, 1, 1, 2009, 4, 1, 24.657534d, 3)]
        [DataRow(2009, 1, 5, 2010, 6, 10, 144.722222d, 2)]
        public void AccrintMtests(int iYear, int iMonth, int iDay, int sYear, int sMonth, int sDay, double expectedResult, int basis)
        {
            var issue = new DateTime(iYear, iMonth, iDay);
            var settlement = new DateTime(sYear, sMonth, sDay);
            _worksheet.Cells["A1"].Value = issue;
            _worksheet.Cells["A2"].Value = settlement;
            _worksheet.Cells["A3"].Value = 0.1;
            _worksheet.Cells["A4"].Value = 1000;
            _worksheet.Cells["A5"].Value = basis;
            _worksheet.Cells["A9"].Formula = "ACCRINTM(A1, A2, A3, A4, A5)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A9"].Value;

            Assert.AreEqual(expectedResult, System.Math.Round((double)result, 6));
        }
    }
}
