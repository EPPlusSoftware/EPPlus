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
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class RefAndLookupTests : FormulaParserTestBase
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("Test");
            _parser = _package.Workbook.FormulaParser;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void VLookupShouldReturnCorrespondingValue()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("test");
                var lookupAddress = "A1:B2";
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Cells["B2"].Value = 5;
                ws.Cells["A3"].Formula = "VLOOKUP(2, " + lookupAddress + ", 2)";
                ws.Calculate();
                var result = ws.Cells["A3"].Value;
                Assert.AreEqual(5, result);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("test");
                var lookupAddress = "A1:B2";
                ws.Cells["A1"].Value = 3;
                ws.Cells["B1"].Value = 1;
                ws.Cells["A2"].Value = 5;
                ws.Cells["B2"].Value = 5;
                ws.Cells["A3"].Formula = "VLOOKUP(4, " + lookupAddress + ", 2, true)";
                ws.Calculate();
                var result = ws.Cells["A3"].Value;
                Assert.AreEqual(1, result);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnCorrespondingValue()
        {
            var lookupAddress = "A1:B2";
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["B1"].Value = 2;
            _worksheet.Cells["A2"].Value = 2;
            _worksheet.Cells["B2"].Value = 5;
            _worksheet.Cells["A3"].Formula = "HLOOKUP(2, " + lookupAddress + ", 2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A3"].Value;
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void HLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
        {
            var lookupAddress = "A1:B2";
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells[1, 1].Value = 3;
                s.Cells[1, 2].Value = 5;
                s.Cells[2, 1].Value = 1;
                s.Cells[2, 2].Value = 2;
                s.Cells[5, 5].Formula = "HLOOKUP(4, " + lookupAddress + ", 2, true)";
                s.Calculate();
                Assert.AreEqual(1, s.Cells[5, 5].Value);
            }
        }

        [TestMethod]
        public void LookupShouldReturnMatchingValue()
        {
            var lookupAddress = "A1:B2";
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells[1, 1].Value = 3;
                s.Cells[1, 2].Value = 5;
                s.Cells[2, 1].Value = 4;
                s.Cells[2, 2].Value = 1;
                s.Cells[5, 5].Formula = "LOOKUP(4, " + lookupAddress + ")";
                s.Calculate();
                Assert.AreEqual(1, s.Cells[5, 5].Value);
            }
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValue()
        {
            var lookupAddress = "A1:A2";

            _worksheet.Cells["A1"].Value = 3;
            _worksheet.Cells["A2"].Value = 5;
            _worksheet.Cells["A3"].Formula = "MATCH(3, " + lookupAddress + ")";
            _worksheet.Calculate();
            Assert.AreEqual(1, _worksheet.Cells["A3"].Value);

        }

        [TestMethod]
        public void RowShouldReturnRowNumber()
        {
            _worksheet.Cells["A4"].Formula = "Row()";
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void RowSholdHandleReference()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "ROW(A4)";
                s1.Calculate();
                Assert.AreEqual(4, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void ColumnShouldReturnRowNumber()
        {
            var ws = _package.Workbook.Worksheets.Add("column");
            ws.Cells["B4"].Formula = "Column()";
            var result = _parser.ParseAt("column",4,2);
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void ColumnSholdHandleReference()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "COLUMN(B4)";
                s1.Calculate();
                Assert.AreEqual(2, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void RowsShouldReturnNbrOfRows()
        {
            _worksheet.Cells["A4"].Formula = "Rows(A5:B7)";
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void ColumnsShouldReturnNbrOfCols()
        {
            using (var package = new ExcelPackage())
            {
                _worksheet.Cells["A4"].Formula = "Columns(A5:B7)";
                
                var result = _parser.ParseAt("A4");
                Assert.AreEqual(2, result);
            }
        }

        [TestMethod]
        public void ChooseShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Choose(1, \"A\", \"B\")");
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void AddressShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Address(1, 1)");
            Assert.AreEqual("$A$1", result);
        }

        [TestMethod]
        public void IndirectShouldReturnARange()
        {
            using (var package = new ExcelPackage(new MemoryStream()))
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["A1:A2"].Value = 2;
                s1.Cells["A3"].Formula = "SUM(Indirect(\"A1:A2\"))";
                s1.Calculate();
                Assert.AreEqual(4d, s1.Cells["A3"].Value);

                s1.Cells["A4"].Formula = "SUM(Indirect(\"A1:A\" & \"2\"))";
                s1.Calculate();
                Assert.AreEqual(4d, s1.Cells["A4"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnASingleValue()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "OFFSET(A1, 2, 1)";
                s1.Calculate();
                Assert.AreEqual(1d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARange()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1))";
                s1.Calculate();
                Assert.AreEqual(3d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARange2()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 10d;
                s1.Cells["B2"].Value = 10d;
                s1.Cells["B3"].Value = 10d;
                s1.Cells["A5"].Formula = "COUNTA(OFFSET(Test!B1, 0, 0, Test!B2, Test!B3))";
                s1.Calculate();
                Assert.AreEqual(3d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetDirectReferenceToMultiRangeShouldSpillValues()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "OFFSET(A1:A3, 0, 1)";
                s1.Calculate();
                var result = s1.Cells["A5"].Value;
                Assert.AreEqual(1d, s1.Cells["A5"].Value);
                Assert.AreEqual(1d, s1.Cells["A6"].Value);
                Assert.AreEqual(1d, s1.Cells["A7"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARangeAccordingToWidth()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2))";
                s1.Calculate();
                Assert.AreEqual(2d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARangeAccordingToHeight()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["C1"].Value = 2d;
                s1.Cells["C2"].Value = 2d;
                s1.Cells["C3"].Value = 2d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2, 2))";
                s1.Calculate();
                Assert.AreEqual(6d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldCoverMultipleColumns()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["C1"].Value = 1d;
                s1.Cells["C2"].Value = 1d;
                s1.Cells["C3"].Value = 1d;
                s1.Cells["D1"].Value = 2d;
                s1.Cells["D2"].Value = 2d;
                s1.Cells["D3"].Value = 2d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:B3, 0, 2))";
                s1.Calculate();
                Assert.AreEqual(9d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void LookupShouldReturnFromResultVector()
        {
            var lookupAddress = "A1:A5";
            var resultAddress = "B1:B5";
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                //lookup_vector
                s.Cells[1, 1].Value = 4.14;
                s.Cells[2, 1].Value = 4.19;
                s.Cells[3, 1].Value = 5.17;
                s.Cells[4, 1].Value = 5.77;
                s.Cells[5, 1].Value = 6.39;
                //result_vector
                s.Cells[1, 2].Value = "red";
                s.Cells[2, 2].Value = "orange";
                s.Cells[3, 2].Value = "yellow";
                s.Cells[4, 2].Value = "green";
                s.Cells[5, 2].Value = "blue";
                //lookup_value
                s.Cells[1, 3].Value = 4.14;
                s.Cells[5, 5].Formula = "LOOKUP(C1, " + lookupAddress + ", " + resultAddress + ")";
                s.Calculate();
                Assert.AreEqual("red", s.Cells[5, 5].Value);
            }
        }

        [TestMethod]
        public void LookupShouldCompareEqualDateWithDouble()
        {
            var date = new DateTime(2020, 2, 7).Date;
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                //lookup_vector
                s.Cells[1, 1].Value = date;
                //result vector
                s.Cells[1, 2].Value = 10;

                //lookup value
                s.Cells[1, 3].Value = date;
                s.Cells[1, 4].Formula = "LOOKUP(C1, A1:A2, B1:B2)";
                s.Calculate();
                Assert.AreEqual(10, s.Cells[1, 4].Value);
            }
        }
        [TestMethod]
        public void OffsetInSecondPartOfRange()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                package.Workbook.FormulaParser.Configure(x => x.AllowCircularReferences = true);
                s.Cells[1, 1].Value = 3;
                s.Cells[2, 1].Value = 5;
                s.Cells[3, 1].Formula = "SUM(A1:OFFSET(A3,-1,0))";
                s.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });
                Assert.AreEqual(8d, s.Cells[3, 1].Value);
            }
        }

        [TestMethod]
        public void OffsetInFirstPartOfRange()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells[1, 1].Value = 3;
                s.Cells[2, 1].Value = 5;
                s.Cells[4, 1].Formula = "SUM(OFFSET(A3,-1,0):A1)";
                s.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });
                Assert.AreEqual(8d, s.Cells[4, 1].Value);
            }
        }

        [TestMethod]
        public void OffsetInBothPartsOfRange()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells[1, 1].Value = 3;
                s.Cells[2, 1].Value = 5;
                s.Cells[4, 1].Formula = "SUM(OFFSET(A3,-2,0):OFFSET(A3,-1,0))";
                s.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });
                Assert.AreEqual(8d, s.Cells[4, 1].Value);
            }
        }
    }
}
