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
    public class TextSplitTests : TestBase
    {
        [TestMethod]
        public void TextSplitTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \" \")";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Josh", sheet.Cells["I3"].Value);
        }

        [TestMethod]
        public void TextSplitMultipleDelimitersTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott-Mats,Jimmy,Cameron-Luther-Josh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, {\"-\",\",\"})";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Josh", sheet.Cells["I3"].Value);
        }

        [TestMethod]
        public void TextSplitMultipleDelimiters2Test()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott-Mats,Jimmy,Cameron-Luther-Josh";
            sheet.Cells["A2"].Value = "-";
            sheet.Cells["A3"].Value = ",";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, A2:A3)";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Josh", sheet.Cells["I3"].Value);
        }

        [TestMethod]
        public void TextSplitMultipleDelimiters3Test()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott-Mats,Jimmy,Cameron-Luther-Josh";
            sheet.Cells["A2"].Value = "-";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, A2:A3)";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D3"].Value);
        }

        [TestMethod]
        public void TextSplitNoDelimiterTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \".\")";
            sheet.Calculate();
            Assert.AreEqual("Scott Mats Jimmy Cameron Luther Josh", sheet.Cells["D3"].Value);
        }

        [TestMethod]
        public void TextSplitRowsAndColumnsTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy\nCameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \" \", \"\n\")";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Cameron", sheet.Cells["D4"].Value);
            Assert.AreEqual("Jimmy", sheet.Cells["F3"].Value);
            Assert.AreEqual("Josh", sheet.Cells["F4"].Value);
        }

        [TestMethod]
        public void TextSplitRowsAndColumnsSwitchedTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy-Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \"-\", \" \")";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Josh", sheet.Cells["D7"].Value);
            Assert.AreEqual("Cameron", sheet.Cells["E5"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["E3"].Value);
        }

        [TestMethod]
        public void TextSplitRowsAndColumnsPaddedTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott Mats Jimmy-Cameron Luther Josh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \"-\", \" \",,,\"Greger\")";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Josh", sheet.Cells["D7"].Value);
            Assert.AreEqual("Cameron", sheet.Cells["E5"].Value);
            Assert.AreEqual("Greger", sheet.Cells["E3"].Value);
            Assert.AreEqual("Greger", sheet.Cells["E7"].Value);
        }

        [TestMethod]
        public void TextSplitIgnoreEmptySetTRUETest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott, Mats,, Jimmy, Cameron,, Luther, Josh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \",\",,TRUE)";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual(" Mats", sheet.Cells["E3"].Value);
            Assert.AreEqual(" Jimmy", sheet.Cells["F3"].Value);
            Assert.AreEqual(" Cameron", sheet.Cells["G3"].Value);
            Assert.AreEqual(" Luther", sheet.Cells["H3"].Value);
        }

        [TestMethod]
        public void TextSplitIgnoreEmptySet1Test()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Scott, Mats,, Jimmy, Cameron,, Luther, Josh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \",\",,1)";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual(" Mats", sheet.Cells["E3"].Value);
            Assert.AreEqual(" Jimmy", sheet.Cells["F3"].Value);
            Assert.AreEqual(" Cameron", sheet.Cells["G3"].Value);
            Assert.AreEqual(" Luther", sheet.Cells["H3"].Value);
        }

        [TestMethod]
        public void TextSplitMatchModeTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "ScottxMatsXJimmyxCameronxLutherXJosh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \"x\")";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("MatsXJimmy", sheet.Cells["E3"].Value);
            Assert.AreEqual("Cameron", sheet.Cells["F3"].Value);
            Assert.AreEqual("LutherXJosh", sheet.Cells["G3"].Value);
        }

        [TestMethod]
        public void TextSplitMatchModeSet1Test()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "ScottxMatsXJimmyxCameronxLutherXJosh";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \"x\",,,1)";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Mats", sheet.Cells["E3"].Value);
            Assert.AreEqual("Jimmy", sheet.Cells["F3"].Value);
            Assert.AreEqual("Cameron", sheet.Cells["G3"].Value);
            Assert.AreEqual("Luther", sheet.Cells["H3"].Value);
            Assert.AreEqual("Josh", sheet.Cells["I3"].Value);
        }

        [TestMethod]
        public void TextSplitFullTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "ScottxMatsXJimmyxxCameron-xLutherXJoshxx";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \"x\",\"-\",1,1,\"Greger\")";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Mats", sheet.Cells["E3"].Value);
            Assert.AreEqual("Jimmy", sheet.Cells["F3"].Value);
            Assert.AreEqual("Cameron", sheet.Cells["G3"].Value);
            Assert.AreEqual("Luther", sheet.Cells["D4"].Value);
            Assert.AreEqual("Josh", sheet.Cells["E4"].Value);
            Assert.AreEqual("Greger", sheet.Cells["F4"].Value);
            Assert.AreEqual("Greger", sheet.Cells["G4"].Value);
        }

        [TestMethod]
        public void TextSplitRangeTest()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            ws.Cells["A2"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            ws.Cells["A3"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            ws.Cells["A5"].Value = "Scott Mats Jimmy Cameron Luther Josh";
            ws.Cells["D15"].Formula = "TEXTSPLIT(A1:A5,\" \")";
            ws.Calculate();
            Assert.AreEqual("Scott", ws.Cells["D15"].Value);
            Assert.AreEqual("Scott", ws.Cells["D16"].Value);
            Assert.AreEqual("Scott", ws.Cells["D17"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["D18"].Value);
            Assert.AreEqual("Scott", ws.Cells["D19"].Value);
        }

        [TestMethod]
        public void TextSplitFull2Test()
        {
            using var package = OpenPackage("TextSplit.xlsx", true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "ScottxMatsXJimmyxxCameron-xLutherXJoshxx";
            sheet.Cells["D3"].Formula = "TEXTSPLIT(A1, \"x\",\"-\",1,1,\"Greger\")";
            sheet.Calculate();
            Assert.AreEqual("Scott", sheet.Cells["D3"].Value);
            Assert.AreEqual("Mats", sheet.Cells["E3"].Value);
            Assert.AreEqual("Jimmy", sheet.Cells["F3"].Value);
            Assert.AreEqual("Cameron", sheet.Cells["G3"].Value);
            Assert.AreEqual("Luther", sheet.Cells["D4"].Value);
            Assert.AreEqual("Josh", sheet.Cells["E4"].Value);
            Assert.AreEqual("Greger", sheet.Cells["F4"].Value);
            Assert.AreEqual("Greger", sheet.Cells["G4"].Value);
            SaveAndCleanup(package);
        }

    }
}
