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
using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using FakeItEasy;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using AddressFunction = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Address;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class RefAndLookupTests
    {
        [TestMethod]
        public void LookupArgumentsShouldSetSearchedValue()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.AreEqual(1, lookupArgs.SearchedValue);
        }

        [TestMethod]
        public void LookupArgumentsShouldSetRangeAddress()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.AreEqual("A:B", lookupArgs.RangeAddress);
        }

        [TestMethod]
        public void LookupArgumentsShouldSetColIndex()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.AreEqual(2, lookupArgs.LookupIndex);
        }

        [TestMethod]
        public void LookupArgumentsShouldSetRangeLookupToTrueAsDefaultValue()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.IsTrue(lookupArgs.RangeLookup);
        }

        [TestMethod]
        public void LookupArgumentsShouldSetRangeLookupToTrueWhenTrueIsSupplied()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2, true);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.IsTrue(lookupArgs.RangeLookup);
        }

        [TestMethod]
        public void VLookupShouldReturnResultFromMatchingRow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(2,A1:B2,2)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(5, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnClosestValueBelowWhenRangeLookupIsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(4,A1:B2,2,true)";
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 5;
                sheet.Cells[2, 2].Value = 4;
                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnClosestStringValueBelowWhenRangeLookupIsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(\"B\",A1:B2,2,true)";
                sheet.Cells[1, 1].Value = "A";
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = "C";
                sheet.Cells[2, 2].Value = 4;
                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void VLookupShouldIgnoreCase()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(\"b\",A1:B2,2,true)";
                sheet.Cells[1, 1].Value = "A";
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = "C";
                sheet.Cells[2, 2].Value = 4;
                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnResultFromMatchingRow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "HLOOKUP(2,A1:B2,2)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(5, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnNaErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsFalse()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "HLOOKUP(2,A1:B2,2,false)";
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();
                var expectedResult = ExcelErrorValue.Create(eErrorType.NA);
                Assert.AreEqual(expectedResult, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "HLOOKUP(1,A1:B2,2,true)";
                sheet.Cells[1, 1].Value = 2;
                sheet.Cells[1, 2].Value = 3;
                sheet.Cells[2, 1].Value = 3;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();
                var expectedResult = ExcelErrorValue.Create(eErrorType.NA);
                Assert.AreEqual(expectedResult, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void LookupShouldReturnResultFromMatchingRowArrayVertical()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "LOOKUP(4,A1:B3,2)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = "A";
                sheet.Cells[2, 1].Value = 3;
                sheet.Cells[2, 2].Value = "B";
                sheet.Cells[3, 1].Value = 5;
                sheet.Cells[3, 2].Value = "C";
                sheet.Calculate();

                Assert.AreEqual("B", sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void LookupShouldReturnResultFromMatchingRowArrayHorizontal()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "LOOKUP(4,A1:C2,2)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 3;
                sheet.Cells[1, 3].Value = 5;
                sheet.Cells[2, 1].Value = "A";
                sheet.Cells[2, 2].Value = "B";
                sheet.Cells[2, 3].Value = "C";
                sheet.Calculate();

                Assert.AreEqual("B", sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void LookupShouldReturnResultFromMatchingSecondArrayHorizontal()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["C1"].Value = 5;
                sheet.Cells["A3"].Value = "A";
                sheet.Cells["B3"].Value = "B";
                sheet.Cells["C3"].Value = "C";

                sheet.Cells["D1"].Formula = "LOOKUP(4, A1:C1, A3:C3)";
                sheet.Calculate();
                var result = sheet.Cells["D1"].Value;
                Assert.AreEqual("B", result);

            }
        }

        [TestMethod]
        public void LookupShouldReturnResultFromMatchingSecondArrayHorizontalWithOffset()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["C1"].Value = 5;
                sheet.Cells["B3"].Value = "A";
                sheet.Cells["C3"].Value = "B";
                sheet.Cells["D3"].Value = "C";

                sheet.Cells["D1"].Formula = "LOOKUP(4, A1:C1, B3:D3)";
                sheet.Calculate();
                var result = sheet.Cells["D1"].Value;
                Assert.AreEqual("B", result);

            }
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeExact()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "Match(3,A1:C1,0)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 3;
                sheet.Cells[1, 3].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(2, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValVertical_MatchTypeExact()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "Match(3,A1:A3,0)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[2, 1].Value = 3;
                sheet.Cells[3, 1].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(2, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestBelow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "Match(4,A1:C1,1)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 3;
                sheet.Cells[1, 3].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(2, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestAbove()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "Match(6,A1:C1,-1)";
                sheet.Cells[1, 1].Value = 10;
                sheet.Cells[1, 2].Value = 8;
                sheet.Cells[1, 3].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(2, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void MatchShouldReturnFirstItemWhenExactMatch_MatchTypeClosestAbove()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "Match(10,A1:C1,-1)";
                sheet.Cells[1, 1].Value = 10;
                sheet.Cells[1, 2].Value = 8;
                sheet.Cells[1, 3].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void MatchShouldHandleAddressOnOtherSheet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells["A1"].Formula = "Match(10, Sheet2!A1:Sheet2!A3, 0)";
                sheet2.Cells["A1"].Value = 9;
                sheet2.Cells["A2"].Value = 10;
                sheet2.Cells["A3"].Value = 11;
                sheet1.Calculate();
                Assert.AreEqual(2, sheet1.Cells["A1"].Value);
            }    
        }

        [TestMethod]
        public void RowShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
        {
            var func = new Row();
            var parsingContext = ParsingContext.Create();
            var rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>(), ParsingContext.Create());
            parsingContext.CurrentCell = new FormulaCellAddress(0, 2, 1);
            var result = func.Execute(Enumerable.Empty<FunctionArgument>(), parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void RowShouldReturnRowSuppliedAddress()
        {
            var func = new Row();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A3"), parsingContext);
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void ColumnShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
        {
            var func = new Column();
            var parsingContext = ParsingContext.Create();
            var rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>(), ParsingContext.Create());
            parsingContext.CurrentCell = new FormulaCellAddress(0, 2, 2);
            var result = func.Execute(Enumerable.Empty<FunctionArgument>(), parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void ColumnShouldReturnRowSuppliedAddress()
        {
            var func = new Column();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("E3"), parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void RowsShouldReturnNbrOfRowsSuppliedRange()
        {
            var func = new Rows();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A1:B3"), parsingContext);
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void RowsShouldReturnNbrOfRowsForEntireColumn()
        {
            var func = new Rows();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A:B"), parsingContext);
            Assert.AreEqual(1048576, result.Result);
        }

        [TestMethod]
        public void ColumnssShouldReturnNbrOfRowsSuppliedRange()
        {
            var func = new Columns();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A1:E3"), parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void ChooseShouldReturnItemByIndex()
        {
            var func = new Choose();
            var parsingContext = ParsingContext.Create();
            var result = func.Execute(FunctionsHelper.CreateArgs(1, "A", "B"), parsingContext);
            Assert.AreEqual("A", result.Result);
        }

        [TestMethod]
        public void AddressShouldReturnAddressByIndexWithDefaultRefType()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2), parsingContext);
            Assert.AreEqual("$B$1", result.Result);
        }

        [TestMethod]
        public void AddressShouldReturnAddressByIndexWithRelativeType()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn), parsingContext);
            Assert.AreEqual("B1", result.Result);
        }

        [TestMethod]
        public void AddressShouldReturnAddressByWithSpecifiedWorksheet()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, true, "Worksheet1"), parsingContext);
            Assert.AreEqual("Worksheet1!B1", result.Result);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void AddressShouldThrowIfR1C1FormatIsSpecified()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, false), parsingContext);
        }
    }
}
