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
    public class CountIfTests
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
        public void CountIfNumeric()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 2d;
            _worksheet.Cells["A3"].Value = 3d;
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, ">1");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfNonNumeric()
        {
            _worksheet.Cells["A1"].Value = "Monday";
            _worksheet.Cells["A2"].Value = "Tuesday";
            _worksheet.Cells["A3"].Value = "Thursday";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "T*day");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        public void CountIfNullExpression()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A3"].Value = null;
            _worksheet.Cells["B2"].Value = null;
            var func = new CountIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 2, 2, 2, 2);
            var args = FunctionsHelper.CreateArgs(range1, range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0d, result.Result);
        }

        [TestMethod]
        public void CountIfNumericExpression()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, 1d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

[TestMethod]
        public void CountIfEqualToEmptyString()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfNotEqualToNull()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "<>");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 0d;
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "0");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfNotEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 0d;
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "<>0");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfGreaterThanZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 1d;
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, ">0");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfGreaterThanOrEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 1d;
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, ">=0");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfLessThanZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = -1d;
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "<0");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfLessThanOrEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = -1d;
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "<=0");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfLessThanCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "<a");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfLessThanOrEqualToCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, "<=a");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfGreaterThanCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, ">a");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfGreaterThanOrEqualToCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, ">=a");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfIgnoreNumericString()
        {
            _worksheet.Cells["A1"].Value = "1";
            _worksheet.Cells["A2"].Value = 2;
            _worksheet.Cells["A3"].Value = "3";
            var func = new CountIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, ">0");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }
    }
}
