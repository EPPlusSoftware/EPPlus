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
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class IndexTests
    {
        private ParsingContext _parsingContext;
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _package = new ExcelPackage(new MemoryStream());
            _worksheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }
        
        [TestMethod]
        public void Index_Should_Return_Value_By_Index()
        {
            var func = new Index();
            var r = new InMemoryRange(new RangeDefinition(3,1));
            r.SetValue(0, 0, 1);
            r.SetValue(1, 0, 2);
            r.SetValue(2, 0, 5);
            var result = func.Execute(
                FunctionsHelper.CreateArgs(
                    r,
                    3
                    ),_parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void Index_Should_Handle_SingleRange()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 3d;
            _worksheet.Cells["A3"].Value = 5d;

            _worksheet.Cells["A4"].Formula = "INDEX(A1:A3,3)";

            _worksheet.Calculate();

            Assert.AreEqual(5d, _worksheet.Cells["A4"].Value);
        }
        [TestMethod]
        public void Index_Should_Handle_SingleRange_Columnwise()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["B1"].Value = 3d;
            _worksheet.Cells["C1"].Value = 5d;

            _worksheet.Cells["A2"].Formula = "INDEX(A1:V1,3)";

            _worksheet.Calculate();

            Assert.AreEqual(5d, _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void Index_SameColumn()
        {
            _worksheet.Cells["A1"].Value = "value";
            _worksheet.Cells["A2"].Formula = "INDEX(A:A,1,1)";

            _worksheet.Calculate();

            Assert.AreEqual("value", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void Index_With_Match()
        {
            _worksheet.Cells["A1"].Value = "key1";
            _worksheet.Cells["A2"].Value = "key2";
            _worksheet.Cells["B1"].Value = "value_to_match";
            _worksheet.Cells["B2"].Formula = "INDEX($B:$B, MATCH(\"key1\", $A:$A, FALSE), 1)";

            _worksheet.Calculate();

            Assert.AreEqual("value_to_match", _worksheet.Cells["B2"].Value);
        }

        [TestMethod]
        public void ShouldHandleSingleValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "asdf";
                sheet.Cells["A2"].Formula = "INDEX(A1,1)";
                sheet.Calculate();
                Assert.AreEqual("asdf", sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void ShouldHandleSingleValue_IndexOutOfRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "asdf";
                sheet.Cells["A2"].Formula = "INDEX(A1,2)";
                sheet.Calculate();
                Assert.AreEqual("#REF!", sheet.Cells["A2"].Value.ToString());
            }
        }
    }
}
