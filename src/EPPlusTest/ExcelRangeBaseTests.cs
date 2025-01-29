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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelRangeBaseTests : TestBase
    {
        [TestMethod]
        public void CopyCopiesCommentsFromSingleCellRanges()
        {
            InitBase();
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
            var sourceExcelRange = ws1.Cells[3, 3];
            Assert.IsNull(sourceExcelRange.Comment);
            sourceExcelRange.AddComment("Testing comment 1", "test1");
            Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
            var destinationExcelRange = ws1.Cells[5, 5];
            Assert.IsNull(destinationExcelRange.Comment);
            sourceExcelRange.Copy(destinationExcelRange);
            // Assert the original comment is intact.
            Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
            // Assert the comment was copied.
            Assert.AreEqual("test1", destinationExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", destinationExcelRange.Comment.Text);
        }

        [TestMethod]
        public void CopyCopiesCommentsFromMultiCellRanges()
        {
            InitBase();
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
            var sourceExcelRangeC3 = ws1.Cells[3, 3];
            var sourceExcelRangeD3 = ws1.Cells[3, 4];
            var sourceExcelRangeE3 = ws1.Cells[3, 5];
            Assert.IsNull(sourceExcelRangeC3.Comment);
            Assert.IsNull(sourceExcelRangeD3.Comment);
            Assert.IsNull(sourceExcelRangeE3.Comment);
            sourceExcelRangeC3.AddComment("Testing comment 1", "test1");
            sourceExcelRangeD3.AddComment("Testing comment 2", "test1");
            sourceExcelRangeE3.AddComment("Testing comment 3", "test1");
            Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
            Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
            Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);
            // Copy the full row to capture each cell at once.
            Assert.IsNull(ws1.Cells[5, 3].Comment);
            Assert.IsNull(ws1.Cells[5, 4].Comment);
            Assert.IsNull(ws1.Cells[5, 5].Comment);
            ws1.Cells["3:3"].Copy(ws1.Cells["5:5"]);
            // Assert the original comments are intact.
            Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
            Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
            Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);
            // Assert the comments were copied.
            var destinationExcelRangeC5 = ws1.Cells[5, 3];
            var destinationExcelRangeD5 = ws1.Cells[5, 4];
            var destinationExcelRangeE5 = ws1.Cells[5, 5];
            Assert.AreEqual("test1", destinationExcelRangeC5.Comment.Author);
            Assert.AreEqual("Testing comment 1", destinationExcelRangeC5.Comment.Text);
            Assert.AreEqual("test1", destinationExcelRangeD5.Comment.Author);
            Assert.AreEqual("Testing comment 2", destinationExcelRangeD5.Comment.Text);
            Assert.AreEqual("test1", destinationExcelRangeE5.Comment.Author);
            Assert.AreEqual("Testing comment 3", destinationExcelRangeE5.Comment.Text);
        }

        [TestMethod]
        public void SettingAddressHandlesMultiAddresses()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                var name = package.Workbook.Names.Add("Test", worksheet.Cells[3, 3]);
                name.Address = "Sheet1!C3";
                name.Address = "Sheet1!D3";
                Assert.IsNull(name.Addresses);
                name.Address = "C3:D3,E3:F3";
                Assert.IsNotNull(name.Addresses);
                name.Address = "Sheet1!C3";
                Assert.IsNull(name.Addresses);
            }
        }

        [TestMethod]
        public void ClearFormulasTest()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].Value = 1;
                worksheet.Cells["A2"].Value = 2;
                worksheet.Cells["A3"].Formula = "SUM(A1:A2)";
                worksheet.Cells["A4"].Formula = "SUM(A1:A2)";
                worksheet.Calculate();
                Assert.AreEqual(3d, worksheet.Cells["A3"].Value);
                Assert.AreEqual(3d, worksheet.Cells["A4"].Value);
                Assert.AreEqual("SUM(A1:A2)", worksheet.Cells["A3"].Formula);
                Assert.AreEqual("SUM(A1:A2)", worksheet.Cells["A4"].Formula);
                worksheet.Cells["A3"].ClearFormulas();
                Assert.AreEqual(3d, worksheet.Cells["A3"].Value);
                Assert.AreEqual(string.Empty, worksheet.Cells["A3"].Formula);
                Assert.AreEqual("SUM(A1:A2)", worksheet.Cells["A4"].Formula);
            }
        }

        [TestMethod]
        public void ClearFormulaValuesTest()
        {
            using(ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].Value = 1;
                worksheet.Cells["A2"].Value = 2;
                worksheet.Cells["A3"].Formula = "SUM(A1:A2)";
                worksheet.Cells["A4"].Formula = "SUM(A1:A2)";
                worksheet.Calculate();
                Assert.AreEqual(3d, worksheet.Cells["A3"].Value);
                Assert.AreEqual(3d, worksheet.Cells["A4"].Value);
                worksheet.Cells["A3"].ClearFormulaValues();
                Assert.IsNull(worksheet.Cells["A3"].Value);
                Assert.AreEqual(3d, worksheet.Cells["A4"].Value);
            }
        }
        [TestMethod]
        public void CheckNaNOnSave()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].Value = double.NaN;
                worksheet.Cells["A2"].Value = 0;
                worksheet.Cells["B1:B2"].Formula = "A1+1";
                object yourNanVariable = double.NaN;
                worksheet.Cells["A1"].Value = yourNanVariable is double d && double.IsNaN(d) ? 0 :  yourNanVariable;

                worksheet.Calculate();
                SaveWorkbook("Nan.xlsx", package);
            }
        }

        [TestMethod]
        public void ExcelRangeBaseShouldReturnCorrectText()
        {
            using (var package = OpenPackage("numfRed.xlsx", true))
            {
                var wb = package.Workbook;
                var aNewWs = wb.Worksheets.Add("NewWs");

                var cell = aNewWs.Cells["A1"];
                cell.Value = -5;
                cell.Style.Numberformat.Format = "[Red]-#";

                var value = aNewWs.Cells["A1"].Value;
                var aText = aNewWs.Cells["A1"].Text;

                Assert.AreEqual(value.ToString(), aText);

                SaveAndCleanup(package);
            }
        }


        [TestMethod]
        public void ExcelRangeBaseShouldReturnCorrectTextAdvanced()
        {
            using (var package = OpenPackage("numfAdvancedCyan.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("NewWs");

                var cell = ws.Cells["A1"];
                cell.Value = -5;
                cell.Style.Numberformat.Format = "[Cyan]-#";

                var cell2 = ws.Cells["A2"];
                //Should be same color as cell1
                //Indicies for Color1,Color2 etc. are 1-based in the string
                cell2.Style.Numberformat.Format = $"[Color{(int)ExcelIndexedColorNamedNumFt.Cyan + 1}]-#";

                var cell3 = ws.Cells["A3"];
                //Should become red
                cell3.Style.Numberformat.Format = $"[Color3]-#";

                ws.Cells["B1:B2"].Value = "Should be same color (Cyan)";
                ws.Cells["B3"].Value = "Should be Red";

                ws.Cells["A2:A3"].Formula = "ROW()";

                ws.Calculate();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ExcelRangeBaseShouldReturnCorrectNumberFormatText()
        {
            using (var package = OpenPackage("numfRedStandard.xlsx", true))
            {
                var wb = package.Workbook;
                var aNewWs = wb.Worksheets.Add("NewWs");

                var cell = aNewWs.Cells["A1"];
                cell.Value = -5;
                cell.Style.Numberformat.Format = "#,##0_);[Red](#,##0)";

                var value = aNewWs.Cells["A1"].Value;

                //Should return "(5)" since that's what excel returns on this format
                var aText = aNewWs.Cells["A1"].Text;

                var numberFormat = cell.Style.Numberformat;

                Assert.AreEqual("(5)", aText);

                SaveAndCleanup(package);
            }
        }

    }
}
