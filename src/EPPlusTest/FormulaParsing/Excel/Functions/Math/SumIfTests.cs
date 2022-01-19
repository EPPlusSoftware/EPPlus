/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class SumIfTests
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
        public void SumIfNumeric()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 2d;
            _worksheet.Cells["A3"].Value = 3d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">1", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(8d, result.Result);
        }

        [TestMethod]
        public void SumIfNonNumeric()
        {
            _worksheet.Cells["A1"].Value = "Monday";
            _worksheet.Cells["A2"].Value = "Tuesday";
            _worksheet.Cells["A3"].Value = "Thursday";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "T*day", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(8d, result.Result);
        }

        [TestMethod]
        public void SumIfShouldIgnoreNumericStrings()
        {
            _worksheet.Cells["A1"].Value = 2;
            _worksheet.Cells["A2"].Value = 1;
            _worksheet.Cells["A3"].Value = "4";
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range1, ">1");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void SumIfNumericExpression()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new SumIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, 1d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SumIfEqualToEmptyString()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SumIfNotEqualToNull()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<>", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(8d, result.Result);
        }

        [TestMethod]
        public void SumIfEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 0d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfNotEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 0d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<>0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void SumIfGreaterThanZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfGreaterThanOrEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">=0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfLessThanZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = -1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfLessThanOrEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = -1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<=0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfLessThanCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void SumIfLessThanOrEqualToCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<=a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void SumIfGreaterThanCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfGreaterThanOrEqualToCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfHandleDates()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfShouldHandleBooleanArg()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = true;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["A2"].Value = false;
                sheet.Cells["B2"].Value = 1;
                sheet.Cells["C1"].Formula = "SUMIF(A1:A2,TRUE,B1:B2)";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void SumIfShouldHandleArrayOfCriterias()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "A";
                sheet.Cells["A2"].Value = "B";
                sheet.Cells["A3"].Value = "A";
                sheet.Cells["A4"].Value = "B";
                sheet.Cells["A5"].Value = "C";
                sheet.Cells["A6"].Value = "B";

                sheet.Cells["B1"].Value = 10;
                sheet.Cells["B2"].Value = 20;
                sheet.Cells["B3"].Value = 10;
                sheet.Cells["B4"].Value = 30;
                sheet.Cells["B5"].Value = 40;
                sheet.Cells["B6"].Value = 10;

                sheet.Cells["A9"].Formula = "SUMIF(A1:A6,{\"A\",\"C\"}, B1:B6)";
                sheet.Calculate();
                Assert.AreEqual(60d, sheet.Cells["A9"].Value);
            }
        }

        [TestMethod]
        public void SumIfShouldHandleRangeWithCriterias()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "A";
                sheet.Cells["A2"].Value = "B";
                sheet.Cells["A3"].Value = "A";
                sheet.Cells["A4"].Value = "B";
                sheet.Cells["A5"].Value = "C";
                sheet.Cells["A6"].Value = "B";

                sheet.Cells["B1"].Value = 10;
                sheet.Cells["B2"].Value = 20;
                sheet.Cells["B3"].Value = 10;
                sheet.Cells["B4"].Value = 30;
                sheet.Cells["B5"].Value = 40;
                sheet.Cells["B6"].Value = 10;

                sheet.Cells["D9"].Value = "A";
                sheet.Cells["E9"].Value = "C";

                sheet.Cells["A9"].Formula = "SUMIF(A1:A6,D9:E9, B1:B6)";
                sheet.Calculate();
                Assert.AreEqual(60d, sheet.Cells["A9"].Value);
            }
        }

        [TestMethod]
        public void SumIfSingleCell()
        {
            _worksheet.Cells["A1"].Value = 20;
            _worksheet.Cells["A2"].Formula = "SUMIF(A1,\">0\")";
            _worksheet.Calculate();

            Assert.AreEqual(20d, _worksheet.Cells["A2"].Value);
        }
    }
}
