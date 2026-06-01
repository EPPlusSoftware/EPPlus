using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class RegexFunctionsTests : TestBase
    {
        [TestMethod]
        public void RegexTest()
        {
            using(var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";
                sheet.Cells["A4"].Value = "Stockholm";
                sheet.Cells["A5"].Value = "Örebro";
                sheet.Cells["A6"].Value = "Linköping";

                sheet.Cells["B1"].Value = "Stockholm";
                sheet.Cells["B2"].Value = "^S";
                sheet.Cells["B3"].Value = "Q[0-9]";
                sheet.Cells["B4"].Value = "202[456]";
                sheet.Cells["B5"].Value = "^[0-9]{5}$";
                sheet.Cells["B6"].Value = "[A-ZÅÄÖ][a-zåäö]+";

                sheet.Cells["D1"].Formula = "REGEXTEST(A1:A6, B1:B6)";
                sheet.Calculate();
                Assert.AreEqual(true, sheet.Cells["D1"].Value);
                Assert.AreEqual(false, sheet.Cells["D2"].Value);
                Assert.AreEqual(false, sheet.Cells["D3"].Value);
                Assert.AreEqual(false, sheet.Cells["D4"].Value);
                Assert.AreEqual(false, sheet.Cells["D5"].Value);
                Assert.AreEqual(true, sheet.Cells["D6"].Value);                
            }
        }

        [TestMethod]
        public void RegexTestMultiplCols()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";
                sheet.Cells["A4"].Value = "Stockholm";
                sheet.Cells["A5"].Value = "Örebro";
                sheet.Cells["A6"].Value = "Linköping";

                sheet.Cells["B1"].Value = "Stockholm";
                sheet.Cells["B2"].Value = "Linköping";
                sheet.Cells["B3"].Value = "Örebro";
                sheet.Cells["B4"].Value = "Stockholm";
                sheet.Cells["B5"].Value = "Örebro";
                sheet.Cells["B6"].Value = "Linköping";

                sheet.Cells["C1"].Value = "Stockholm";
                sheet.Cells["C2"].Value = "^S";
                sheet.Cells["C3"].Value = "Q[0-9]";
                sheet.Cells["C4"].Value = "202[456]";
                sheet.Cells["C5"].Value = "^[0-9]{5}$";
                sheet.Cells["C6"].Value = "[A-ZÅÄÖ][a-zåäö]+";

                sheet.Cells["D1"].Formula = "REGEXTEST(A1:B6, C1:C6)";
                sheet.Calculate();
                Assert.AreEqual(true, sheet.Cells["D1"].Value);
                Assert.AreEqual(false, sheet.Cells["D2"].Value);
                Assert.AreEqual(false, sheet.Cells["D3"].Value);
                Assert.AreEqual(false, sheet.Cells["D4"].Value);
                Assert.AreEqual(false, sheet.Cells["D5"].Value);
                Assert.AreEqual(true, sheet.Cells["D6"].Value);

                Assert.AreEqual(true, sheet.Cells["E1"].Value);
                Assert.AreEqual(false, sheet.Cells["E2"].Value);
                Assert.AreEqual(false, sheet.Cells["E3"].Value);
                Assert.AreEqual(false, sheet.Cells["E4"].Value);
                Assert.AreEqual(false, sheet.Cells["E5"].Value);
                Assert.AreEqual(true, sheet.Cells["E6"].Value);
            }
        }

        [TestMethod]
        public void RegexUnevenInputRanges()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";
                sheet.Cells["A4"].Value = "Stockholm";
                sheet.Cells["A5"].Value = "Örebro";
                sheet.Cells["A6"].Value = "Linköping";

                sheet.Cells["B1"].Value = "Stockholm";
                sheet.Cells["B2"].Value = "Linköping";
                sheet.Cells["B3"].Value = "Örebro";
                sheet.Cells["B4"].Value = "Stockholm";
                sheet.Cells["B5"].Value = "Örebro";
                sheet.Cells["B6"].Value = "Linköping";

                sheet.Cells["C1"].Value = 2026;
                sheet.Cells["C2"].Value = 2026;
                sheet.Cells["C3"].Value = 2025;
                sheet.Cells["C4"].Value = 2025;
                sheet.Cells["C5"].Value = 2025;
                sheet.Cells["C6"].Value = 2024;

                sheet.Cells["D4"].Value = "202[456]";
                sheet.Cells["D5"].Value = "^[0-9]{5}$";
                sheet.Cells["D6"].Value = "[A-ZÅÄÖ][a-zåäö]+";
                sheet.Cells["D7"].Value = "[0-9]+";

                sheet.Cells["E1"].Formula = "REGEXTEST(A1:C6, D4:D7)";
                sheet.Calculate();

                Assert.AreEqual(false, sheet.Cells["E1"].Value);
                Assert.AreEqual(false, sheet.Cells["E2"].Value);
                Assert.AreEqual(true, sheet.Cells["E3"].Value);
                Assert.AreEqual(false, sheet.Cells["E4"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["E5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["E6"].Value);

                Assert.AreEqual(false, sheet.Cells["F1"].Value);
                Assert.AreEqual(false, sheet.Cells["F2"].Value);
                Assert.AreEqual(true, sheet.Cells["F3"].Value);
                Assert.AreEqual(false, sheet.Cells["F4"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["F5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["F6"].Value);

                Assert.AreEqual(true, sheet.Cells["G1"].Value);
                Assert.AreEqual(false, sheet.Cells["G2"].Value);
                Assert.AreEqual(false, sheet.Cells["G3"].Value);
                Assert.AreEqual(true, sheet.Cells["G4"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G6"].Value);
            }
        }

        [TestMethod]
        public void RegexTestCaseSensitive()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";
                sheet.Cells["A4"].Value = "Stockholm";
                sheet.Cells["A5"].Value = "Örebro";
                sheet.Cells["A6"].Value = "Linköping";

                sheet.Cells["B1"].Value = "k";

                sheet.Cells["D1"].Formula = "REGEXTEST(A1:A6, B1, 1)";
                sheet.Calculate();

                Assert.AreEqual(true, sheet.Cells["D1"].Value);
                Assert.AreEqual(true, sheet.Cells["D2"].Value);
                Assert.AreEqual(false, sheet.Cells["D3"].Value);
                Assert.AreEqual(true, sheet.Cells["D4"].Value);
                Assert.AreEqual(false, sheet.Cells["D5"].Value);
                Assert.AreEqual(true, sheet.Cells["D6"].Value);
            }
        }

        [TestMethod]
        public void RegexTestCaseSensitiveError()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";
                sheet.Cells["A4"].Value = "Stockholm";
                sheet.Cells["A5"].Value = "Örebro";
                sheet.Cells["A6"].Value = "Linköping";

                sheet.Cells["B1"].Value = "k";

                sheet.Cells["D1"].Formula = "REGEXTEST(A1:A6, B1, 2)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D3"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D4"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D6"].Value);
            }
        }

        [TestMethod]
        public void RegexExtract()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Kossa Mail@mail.se";
                sheet.Cells["A2"].Value = "Får enmail@mef.se sd";
                sheet.Cells["A3"].Value = "mailens@hemma.com";
                sheet.Cells["A4"].Value = "mail@se.se";
                sheet.Cells["A5"].Value = "Tupp ska gala gmail@adress.net dwqdw";
                sheet.Cells["A6"].Value = "Katt";

                sheet.Cells["B1"].Value = "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,}";


                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A6, B1)";
                sheet.Calculate();
                Assert.AreEqual("Mail@mail.se", sheet.Cells["D1"].Value);
                Assert.AreEqual("enmail@mef.se sd", sheet.Cells["D2"].Value);
                Assert.AreEqual("mailens@hemma.com", sheet.Cells["D3"].Value);
                Assert.AreEqual("mail@se.se", sheet.Cells["D4"].Value);
                Assert.AreEqual("gmail@adress.net dwqdw", sheet.Cells["D5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D6"].Value);
            }
        }

        [TestMethod]
        public void RegexExtractReturnMode1()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Just #fitness finished 5k! #running";

                sheet.Cells["B1"].Value = "#\\w+";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1, B1, 1)";
                sheet.Calculate();
                Assert.AreEqual("#fitness", sheet.Cells["D1"].Value);
                Assert.AreEqual("#running", sheet.Cells["E1"].Value);
            }
        }

        [TestMethod]
        public void RegexExtractShouldReturnSingleWithReturnMode1()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Just #fitness finished 5k! #running";
                sheet.Cells["A2"].Value = "Look at this picture #nature #instagram";
                sheet.Cells["B1"].Value = "#\\w+";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A2, B1, 1)";
                sheet.Calculate();
                Assert.AreEqual("#fitness", sheet.Cells["D1"].Value);
                Assert.AreNotEqual("#running", sheet.Cells["E1"].Value);
                Assert.AreEqual("#nature", sheet.Cells["D2"].Value);
                Assert.AreNotEqual("#instagram", sheet.Cells["E2"].Value);
            }
        }


        [TestMethod]
        public void RegexExtractReturnMode2()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "9183-Green-M";

                sheet.Cells["B1"].Value = "(\\d{4})-(\\w+)-(\\w+)";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1, B1, 2)";
                sheet.Calculate();
                Assert.AreEqual("9183", sheet.Cells["D1"].Value);
                Assert.AreEqual("Green", sheet.Cells["E1"].Value);
                Assert.AreEqual("M", sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void RegexExtractShouldReturnSingleWithReturnMode2()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "9183-Green-M";
                sheet.Cells["A2"].Value = "2546-Black-XL";

                sheet.Cells["B1"].Value = "(\\d{4})-(\\w+)-(\\w+)";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A2, B1, 2)";
                sheet.Calculate();

                Assert.AreEqual("9183", sheet.Cells["D1"].Value);                
                Assert.AreEqual("2546", sheet.Cells["D2"].Value);
            }
        }
    }
}
