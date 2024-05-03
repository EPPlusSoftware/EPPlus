/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class TextAfterTests : TestBase
    {
        [TestMethod]
        public void TextAfterTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTAFTER(A1, \" \")";
            sheet.Cells["A2"].Value = "Scott Mats Jimmy-Cameron Luther Josh";
            sheet.Cells["D4"].Formula = "TEXTAFTER(A2, \"-\")";
            sheet.Calculate();
            Assert.AreEqual("Mats Jimmy Cameron Luther Josh", sheet.Cells["D3"].Value);
            Assert.AreEqual("Cameron Luther Josh", sheet.Cells["D4"].Value);
        }

        [TestMethod]
        public void TextAfterInstanceNumTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTAFTER(A1, \" \", 2)";
            sheet.Cells["D4"].Formula = "TEXTAFTER(A1, \" \", 5)";
            sheet.Cells["D5"].Formula = "TEXTAFTER(A1, \" \", 7)";
            sheet.Calculate();
            Assert.AreEqual("Jimmy Cameron Luther Josh", sheet.Cells["D3"].Value);
            Assert.AreEqual("Josh", sheet.Cells["D4"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D5"].Value);
        }

        [TestMethod]
        public void TextAfterInstanceNegativeNumTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTAFTER(A1, \" \", -2)";
            sheet.Cells["D4"].Formula = "TEXTAFTER(A1, \" \", -5)";
            sheet.Cells["D5"].Formula = "TEXTAFTER(A1, \" \", -7)";
            sheet.Calculate();
            Assert.AreEqual("Luther Josh", sheet.Cells["D3"].Value);
            Assert.AreEqual("Mats Jimmy Cameron Luther Josh", sheet.Cells["D4"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D5"].Value);
        }

        [TestMethod]
        public void TextAfterMatchModeTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "ScottXMatsxJimmyxCameronXLutherxJosh";
            sheet.Cells["D3"].Formula = "TEXTAFTER(A1, \"x\", 4, 1)";
            sheet.Calculate();
            Assert.AreEqual("LutherxJosh", sheet.Cells["D3"].Value);
        }

        [TestMethod]
        public void TextAfterMatchEndTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTAFTER(A1, \" \", 6,, 1)";
            sheet.Cells["D4"].Formula = "TEXTAFTER(A1, \" \", -6,, 1)";
            sheet.Cells["D5"].Formula = "TEXTAFTER(A1, \" \", -2,, 1)";
            sheet.Cells["D6"].Formula = "TEXTAFTER(A1, \" \", 2,, 1)";
            sheet.Cells["D7"].Formula = "TEXTAFTER(A1, \" \", 7,, 1)";
            sheet.Cells["D8"].Formula = "TEXTAFTER(A1, \" \", -7,, 1)";
            sheet.Calculate();
            Assert.AreEqual("Scott Mats Jimmy Cameron Luther Josh", sheet.Cells["D3"].Value);
            Assert.AreEqual("Scott Mats Jimmy Cameron Luther Josh", sheet.Cells["D4"].Value);
            Assert.AreEqual("Luther Josh", sheet.Cells["D5"].Value);
            Assert.AreEqual("Jimmy Cameron Luther Josh", sheet.Cells["D6"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D7"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D8"].Value);
        }

        [TestMethod]
        public void TextAfterIfNotFoundTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTAFTER(A1, \".\",,,,\"Test\")";
            sheet.Cells["D4"].Formula = "TEXTAFTER(A1, \".\",7,,,\"Test\")";
            sheet.Cells["D5"].Formula = "TEXTAFTER(A1, \".\",-8,,,\"Test\")";
            sheet.Cells["D6"].Formula = "TEXTAFTER(A1, \".\",7,,1,\"Test\")";
            sheet.Calculate();
            Assert.AreEqual("Test", sheet.Cells["D3"].Value);
            Assert.AreEqual("Test", sheet.Cells["D4"].Value);
            Assert.AreEqual("Test", sheet.Cells["D5"].Value);
            Assert.AreEqual("Test", sheet.Cells["D6"].Value);
        }

        [TestMethod]
        public void TextAfterMultipleDelimitersTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott,Mats-Jimmy-Cameron,Luther,Josh";
            sheet.Cells["D3"].Formula = "TEXTAFTER(A1, {\",\", \"-\"}, -4)";
            sheet.Calculate();
            Assert.AreEqual("Jimmy-Cameron,Luther,Josh", sheet.Cells["D3"].Value);
        }

        [TestMethod]
        public void TextAfterRangeText()
        {
            using var package = OpenTemplatePackage("TextAfterTest.xlsx");
            var sheet = package.Workbook.Worksheets["Sheet1"];
            sheet.Cells["A4"].Value = "Scott Mats Jimmy";
            sheet.Cells["A5"].Value = "Cameron Luther Josh";
            sheet.Cells["B4"].Value = "Cameron Luther Josh";
            sheet.Cells["D12"].Formula = "TEXTAFTER(A4:A5, \" \")";
            sheet.Cells["E12"].Formula = "TEXTAFTER(A4:B4, \" \")";
            sheet.Calculate();
            Assert.AreEqual("Mats Jimmy", sheet.Cells["D12"].Value);
            Assert.AreEqual("Luther Josh", sheet.Cells["D13"].Value);
            Assert.AreEqual("Mats Jimmy", sheet.Cells["E12"].Value);
            Assert.AreEqual("Luther Josh", sheet.Cells["F12"].Value);
            SaveAndCleanup(package);
        }

        [TestMethod]
        public void TextAfterCreateWorkBookTest()
        {
            using var package = OpenPackage("TextAfter.xlsx", true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["A2"].Value = "ScottXMatsxJimmyxCameronXLutherxJosh";
            sheet.Cells["A3"].Value = "Scott,Mats-Jimmy-Cameron,Luther,Josh";
            sheet.Cells["D3"].Formula = "TEXTAFTER(A1, \" \")";
            sheet.Cells["D4"].Formula = "TEXTAFTER(A1, \" \", 2)";
            sheet.Cells["D5"].Formula = "TEXTAFTER(A1, \" \", -2)";
            sheet.Cells["D6"].Formula = "TEXTAFTER(A2, \"x\", 4, 1)";
            sheet.Cells["D7"].Formula = "TEXTAFTER(A1, \" \", 2,, 1)";
            sheet.Cells["D8"].Formula = "TEXTAFTER(A1, \" \", 7,, 1)";
            sheet.Cells["D9"].Formula = "TEXTAFTER(A1, \".\",,,,\"Test\")";
            sheet.Cells["D10"].Formula = "TEXTAFTER(A3, {\",\", \"-\"}, 4)";
            sheet.Calculate();
            Assert.AreEqual("Mats Jimmy Cameron Luther Josh", sheet.Cells["D3"].Value);
            Assert.AreEqual("Jimmy Cameron Luther Josh", sheet.Cells["D4"].Value);
            Assert.AreEqual("Luther Josh", sheet.Cells["D5"].Value);
            Assert.AreEqual("LutherxJosh", sheet.Cells["D6"].Value);
            Assert.AreEqual("Jimmy Cameron Luther Josh", sheet.Cells["D7"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D8"].Value);
            Assert.AreEqual("Test", sheet.Cells["D9"].Value);
            Assert.AreEqual("Luther,Josh", sheet.Cells["D10"].Value);
            SaveAndCleanup(package);
        }
    }
}
