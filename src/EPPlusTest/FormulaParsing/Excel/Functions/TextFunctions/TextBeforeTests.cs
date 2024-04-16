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
    public class TextBeforeTests : TestBase
    {
        [TestMethod]
        public void TextBeforeTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTBEFORE(A1, \" \")";
            sheet.Cells["A2"].Value = "Scott Mats Jimmy-Cameron Luther Josh";
            sheet.Cells["D4"].Formula = "TEXTBEFORE(A2, \"-\")";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Scott Mats Jimmy", sheet.Cells["D4"].Value);
        }

        [TestMethod]
        public void TextBeforeInstanceNumTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTBEFORE(A1, \" \", 2)";
            sheet.Cells["D4"].Formula = "TEXTBEFORE(A1, \" \", 5)";
            sheet.Cells["D5"].Formula = "TEXTBEFORE(A1, \" \", 7)";
            sheet.Calculate();
            Assert.AreEqual("Scott Mats", sheet.Cells["D3"].Value);
            Assert.AreEqual("Scott Mats Jimmy Cameron Luther", sheet.Cells["D4"].Value);
            Assert.AreEqual(ExcelErrorValue.Values.NA, sheet.Cells["D5"].Value.ToString());
        }

        [TestMethod]
        public void TextBeforeInstanceNegativeNumTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTBEFORE(A1, \" \", -2)";
            sheet.Cells["D4"].Formula = "TEXTBEFORE(A1, \" \", -5)";
            sheet.Cells["D5"].Formula = "TEXTBEFORE(A1, \" \", -7)";
            sheet.Calculate();
            Assert.AreEqual("Scott Mats Jimmy Cameron", sheet.Cells["D3"].Value);
            Assert.AreEqual("Scott", sheet.Cells["D4"].Value);
            Assert.AreEqual(ExcelErrorValue.Values.NA, sheet.Cells["D5"].Value.ToString());
        }

        [TestMethod]
        public void TextBeforeMatchModeTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "ScottXMatsxJimmyxCameronXLutherxJosh";
            sheet.Cells["D3"].Formula = "TEXTBEFORE(A1, \"x\", 4, 1)";
            sheet.Calculate();
            Assert.AreEqual("ScottXMatsxJimmyxCameron", sheet.Cells["D3"].Value);
        }

        [TestMethod]
        public void TextBeforeMatchEndTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTBEFORE(A1, \" \", 6,, 1)";
            sheet.Cells["D4"].Formula = "TEXTBEFORE(A1, \" \", -6,, 1)";
            sheet.Cells["D5"].Formula = "TEXTBEFORE(A1, \" \", -2,, 1)";
            sheet.Cells["D6"].Formula = "TEXTBEFORE(A1, \" \", 2,, 1)";
            sheet.Cells["D7"].Formula = "TEXTBEFORE(A1, \" \", 7,, 1)";
            sheet.Cells["D8"].Formula = "TEXTBEFORE(A1, \" \", 7,, 1)";
            sheet.Calculate();
            Assert.AreEqual("Scott Mats Jimmy Cameron Luther Josh", sheet.Cells["D3"].Value);
            Assert.AreEqual("Scott Mats Jimmy Cameron Luther Josh", sheet.Cells["D4"].Value);
            Assert.AreEqual("Scott Mats Jimmy Cameron", sheet.Cells["D5"].Value);
            Assert.AreEqual("Scott Mats", sheet.Cells["D6"].Value);
            Assert.AreEqual(ExcelErrorValue.Values.NA, sheet.Cells["D7"].Value.ToString());
            Assert.AreEqual(ExcelErrorValue.Values.NA, sheet.Cells["D8"].Value.ToString());
        }

        [TestMethod]
        public void TextBeforeIfNotFoundTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTBEFORE(A1, \".\",,,,\"Test\")";
            sheet.Cells["D4"].Formula = "TEXTBEFORE(A1, \".\",7,,,\"Test\")";
            sheet.Cells["D5"].Formula = "TEXTBEFORE(A1, \".\",-8,,,\"Test\")";
            sheet.Cells["D6"].Formula = "TEXTBEFORE(A1, \".\",7,,1,\"Test\")";
            sheet.Calculate();
            Assert.AreEqual("Test", sheet.Cells["D3"].Value);
            Assert.AreEqual("Test", sheet.Cells["D4"].Value);
            Assert.AreEqual("Test", sheet.Cells["D5"].Value);
            Assert.AreEqual("Test", sheet.Cells["D6"].Value);
        }

        [TestMethod]
        public void TextBeforeMultipleDelimitersTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott,Mats-Jimmy-Cameron,Luther,Josh";
            sheet.Cells["D3"].Formula = "TEXTBEFORE(A1, {\",\", \"-\"}, 4)";
            sheet.Calculate();
            Assert.AreEqual("Scott,Mats-Jimmy-Cameron", sheet.Cells["D3"].Value);
        }

        [TestMethod]
        public void TextBeforeCreateWorkBookTest()
        {
            using var package = OpenPackage("TextBefore.xlsx", true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["A2"].Value = "ScottXMatsxJimmyxCameronXLutherxJosh";
            sheet.Cells["A3"].Value = "Scott,Mats-Jimmy-Cameron,Luther,Josh";
            sheet.Cells["D3"].Formula = "TEXTBEFORE(A1, \" \")";
            sheet.Cells["D4"].Formula = "TEXTBEFORE(A1, \" \", 2)";
            sheet.Cells["D5"].Formula = "TEXTBEFORE(A1, \" \", -2)";
            sheet.Cells["D6"].Formula = "TEXTBEFORE(A2, \"x\", 4, 1)";
            sheet.Cells["D7"].Formula = "TEXTBEFORE(A1, \" \", 2,, 1)";
            sheet.Cells["D8"].Formula = "TEXTBEFORE(A1, \" \", 7,, 1)";
            sheet.Cells["D9"].Formula = "TEXTBEFORE(A1, \".\",,,,\"Test\")";
            sheet.Cells["D10"].Formula = "TEXTBEFORE(A3, {\",\", \"-\"}, 4)";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Scott Mats", sheet.Cells["D4"].Value);
            Assert.AreEqual("Scott Mats Jimmy Cameron", sheet.Cells["D5"].Value);
            Assert.AreEqual("ScottXMatsxJimmyxCameron", sheet.Cells["D6"].Value);
            Assert.AreEqual("Scott Mats", sheet.Cells["D7"].Value);
            Assert.AreEqual(ExcelErrorValue.Values.NA, sheet.Cells["D8"].Value.ToString());
            Assert.AreEqual("Test", sheet.Cells["D9"].Value);
            Assert.AreEqual("Scott,Mats-Jimmy-Cameron", sheet.Cells["D10"].Value);
            SaveAndCleanup(package);
        }
    }
}
