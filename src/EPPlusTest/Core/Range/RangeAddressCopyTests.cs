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
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System.Drawing;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeAddressCopyTests : TestBase
    {
        [TestMethod]
        public void ValidateCopyFormulasRow()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("CopyRowWise");
                ws.Cells["A1:C3"].Value = 1;
                ws.Cells["D3"].Formula = "A1";
                ws.Cells["E3"].Formula = "B2";
                ws.Cells["F3"].Formula = "C3";
                ws.Cells["G3"].Formula = "A$1";
                ws.Cells["H3"].Formula = "B$2";
                ws.Cells["J3"].Formula = "C$3";

                //Validate that formulas are copied correctly row-wise
                ws.Cells["D3"].Copy(ws.Cells["D2"]);
                Assert.AreEqual("#REF!", ws.Cells["D2"].Formula);
                ws.Cells["E3"].Copy(ws.Cells["E2"]);
                Assert.AreEqual("B1", ws.Cells["E2"].Formula);
                ws.Cells["F3"].Copy(ws.Cells["F2"]);
                Assert.AreEqual("C2", ws.Cells["F2"].Formula);
                ws.Cells["G3"].Copy(ws.Cells["G2"]);
                Assert.AreEqual("A$1", ws.Cells["G2"].Formula);
                ws.Cells["H3"].Copy(ws.Cells["H2"]);
                Assert.AreEqual("B$2", ws.Cells["H2"].Formula);
                ws.Cells["J3"].Copy(ws.Cells["J2"]);
                Assert.AreEqual("C$3", ws.Cells["J2"].Formula);

                ws.Cells["D3"].Copy(ws.Cells["D1"]);
                Assert.AreEqual("#REF!", ws.Cells["D1"].Formula);
                ws.Cells["E3"].Copy(ws.Cells["E1"]);
                Assert.AreEqual("#REF!", ws.Cells["E1"].Formula);
                ws.Cells["F3"].Copy(ws.Cells["F1"]);
                Assert.AreEqual("C1", ws.Cells["F1"].Formula);
                ws.Cells["G3"].Copy(ws.Cells["G1"]);
                Assert.AreEqual("A$1", ws.Cells["G1"].Formula);
                ws.Cells["H3"].Copy(ws.Cells["H1"]);
                Assert.AreEqual("B$2", ws.Cells["H1"].Formula);
                ws.Cells["J3"].Copy(ws.Cells["J1"]);
                Assert.AreEqual("C$3", ws.Cells["J1"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellRow()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM(A1:B1)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM(#REF!)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM(#REF!)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM(B1:C1)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM(A2:B2)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM(#REF!)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM(#REF!)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM(B1:C1)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM(A2:B2)", ws2.Cells["D4"].Formula);

            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellFullColumn()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM(A:A)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM(A:A)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM(#REF!)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM(B:B)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM(A:A)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM(A:A)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM(#REF!)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM(B:B)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM(A:A)", ws2.Cells["D4"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellFullColumnFixed()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM($A:$A)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM($A:$A)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM($A:$A)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM($A:$A)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM($A:$A)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM($A:$A)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM($A:$A)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM($A:$A)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM($A:$A)", ws2.Cells["D4"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellFullRow()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM(1:1)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM(#REF!)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM(1:1)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM(1:1)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM(2:2)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM(#REF!)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM(1:1)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM(1:1)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM(2:2)", ws2.Cells["D4"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellFullRowFixed()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM($1:$1)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM($1:$1)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM($1:$1)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM($1:$1)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM($1:$1)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM($1:$1)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM($1:$1)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM($1:$1)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM($1:$1)", ws2.Cells["D4"].Formula);
            }
        }
        [TestMethod]
        public void Copy_Formula_From_Other_Workbook_Issue_Test()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sourceWs = workbook.Worksheets.Add("Sheet2");
                sourceWs.Cells["A1"].Value = 24;
                sourceWs.Cells["A2"].Value = 75;
                sourceWs.Cells["A3"].Value = 94;
                sourceWs.Cells["A4"].Value = 34;

                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "VLOOKUP($B$1,Sheet2!A:B,2,FALSE)";

                //Literal formula copy to cell A2 - PASSES
                ws.Cells["A1"].Copy(ws.Cells[2, 1], ExcelRangeCopyOptionFlags.ExcludeFormulas);
                ws.Cells["A2"].Formula = ws.Cells[1, 1].Formula;

                Assert.IsFalse(
                    string.IsNullOrWhiteSpace(ws.Cells["A2"].Formula)
                    , "A2 formula should be set"
                );

                Assert.AreEqual(
                    ws.Cells["A1"].Formula
                    , ws.Cells["A2"].Formula
                    , $"{ws.Cells["A2"].Formula} != {ws.Cells["A1"].Formula}"
                );

                //Cell copy to cell A3 - FAILS
                ws.Cells["A1"].Copy(ws.Cells["A3"]);

                Assert.IsFalse(
                    string.IsNullOrWhiteSpace(ws.Cells["A3"].Formula)
                    , "A3 formula should be set"
                );

                Assert.AreEqual("VLOOKUP($B$1,Sheet2!A:B,2,FALSE)", ws.Cells["A3"].Formula);
            }
        }
        [TestMethod]
        public void CopyValuesOnly()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);
                ws.Cells["B5"].Style.Numberformat.Format = "0";
                ws.Cells["A1:A2"].Copy(ws.Cells["B5:B6"], ExcelRangeCopyOptionFlags.ExcludeFormulas, ExcelRangeCopyOptionFlags.ExcludeStyles);

                Assert.AreEqual(1, ws.Cells["B5"].Value);
                Assert.AreEqual(2D, ws.Cells["B6"].Value);

                Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["B6"].Formula));
                Assert.IsFalse(ws.Cells["B5"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Italic);
                Assert.AreEqual("0", ws.Cells["B5"].Style.Numberformat.Format);
            }
        }
        [TestMethod]
        public void CopyStylesOnly()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);
                ws.Cells["B5"].Value = 5;
                ws.Cells["B6"].Value = 7;
                ws.Cells["A1:A2"].Copy(ws.Cells["B5:B6"], ExcelRangeCopyOptionFlags.ExcludeValues);

                Assert.AreEqual(5, ws.Cells["B5"].Value);
                Assert.AreEqual(7, ws.Cells["B6"].Value);
                Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["B6"].Formula));
                Assert.IsTrue(ws.Cells["B5"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["B6"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["B6"].Style.Font.Italic);
            }
        }
        [TestMethod]
        public void CopyDataValidationsSameWorksheet()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);
                var dv = ws.Cells["B2:D5"].DataValidation.AddIntegerDataValidation();
                dv.Formula.Value = 1;
                dv.Formula2.Value = 3;
                dv.ShowErrorMessage = true;
                dv.ErrorStyle = OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle.stop;
                ws.Cells["A1:C4"].Copy(ws.Cells["E5"]);

                Assert.AreEqual("B2:D5,F6:G8", dv.Address.Address);
            }
        }
        [TestMethod]
        public void CopyDataValidationsNewWorksheet()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws1 = SetupCopyRange(p);
                var dv = ws1.Cells["B2:D5"].DataValidation.AddIntegerDataValidation();
                dv.Formula.Value = 1;
                dv.Formula2.Value = 3;
                dv.ShowErrorMessage = true;
                dv.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");
                ws1.Cells["A1:C4"].Copy(ws2.Cells["E5"]);

                Assert.AreEqual(1, ws2.DataValidations.Count);
                var dv2 = ws2.DataValidations[0].As.IntegerValidation;
                Assert.AreEqual("F6:G8", dv2.Address.Address);
                Assert.AreEqual(1, dv2.Formula.Value);
                Assert.AreEqual(3, dv2.Formula2.Value);
                Assert.IsTrue(dv.ShowErrorMessage.Value);
                Assert.AreEqual(ExcelDataValidationWarningStyle.stop, dv.ErrorStyle);

                SaveWorkbook("dvcopy.xlsx", p);
            }
        }
        [TestMethod]
        public void CopyDataValidationNewPackage()
        {
            using (var p1 = new ExcelPackage())
            {
                ExcelWorksheet ws1 = SetupCopyRange(p1);
                var dv = ws1.Cells["B2:D5"].DataValidation.AddIntegerDataValidation();
                dv.Formula.Value = 1;
                dv.Formula2.Value = 3;
                dv.ShowErrorMessage = true;
                dv.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                using (var p2 = new ExcelPackage())
                {
                    var ws2 = p2.Workbook.Worksheets.Add("Sheet Copy");
                    ws1.Cells["A1:C4"].Copy(ws2.Cells["E5"]);

                    Assert.AreEqual(1, ws2.DataValidations.Count);
                    var dv2 = ws2.DataValidations[0].As.IntegerValidation;
                    Assert.AreEqual("F6:G8", dv2.Address.Address);
                    Assert.AreEqual(1, dv2.Formula.Value);
                    Assert.AreEqual(3, dv2.Formula2.Value);
                    Assert.IsTrue(dv.ShowErrorMessage.Value);
                    Assert.AreEqual(ExcelDataValidationWarningStyle.stop, dv.ErrorStyle);

                    SaveWorkbook("dvcopy.xlsx", p2);
                }
            }
        }
        [TestMethod]
        public void CopyConditionalFormattingSameWorkbook()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);
                var cf1 = ws.Cells["B2:D5"].ConditionalFormatting.AddBetween();

                ws.Cells["A1:C4"].Copy(ws.Cells["E5"]);

                Assert.AreEqual("B2:D5,F6:G8", cf1.Address.Address);
            }
        }

        [TestMethod]
        public void CopyConditionalFormattingNewWorksheet()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws1 = SetupCopyRange(p);
                var cf1 = ws1.Cells["B2:D5"].ConditionalFormatting.AddBetween();
                cf1.Formula = "1";
                cf1.Formula2 = "3";
                cf1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cf1.Style.Fill.BackgroundColor.SetColor(Color.Red);
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");
                ws1.Cells["A1:C4"].Copy(ws2.Cells["E5"]);

                Assert.AreEqual(1, ws2.ConditionalFormatting.Count);
                var cf2 = ws2.ConditionalFormatting[0].As.Between;
                Assert.AreEqual("F6:G8", cf2.Address.Address);
                Assert.AreEqual("1", cf2.Formula);
                Assert.AreEqual("3", cf2.Formula2);
                Assert.AreEqual(ExcelFillStyle.Solid, cf2.Style.Fill.PatternType);
                Assert.AreEqual(Color.Red.ToArgb(), cf2.Style.Fill.BackgroundColor.Color.Value.ToArgb());
            }
        }
        [TestMethod]
        public void CopyConditionalFormattingNewPackage()
        {
            using (var p1 = new ExcelPackage())
            {
                ExcelWorksheet ws1 = SetupCopyRange(p1);
                var cf1 = ws1.Cells["B2:D5"].ConditionalFormatting.AddBetween();
                cf1.Formula = "1";
                cf1.Formula2 = "3";
                cf1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cf1.Style.Fill.BackgroundColor.SetColor(Color.Red);
                using (var p2 = new ExcelPackage())
                {
                    var ws2 = p2.Workbook.Worksheets.Add("Sheet2");
                    ws1.Cells["A1:C4"].Copy(ws2.Cells["E5"]);

                    Assert.AreEqual(1, ws2.ConditionalFormatting.Count);
                    var cf2 = ws2.ConditionalFormatting[0].As.Between;
                    Assert.AreEqual("F6:G8", cf2.Address.Address);
                    Assert.AreEqual("1", cf2.Formula);
                    Assert.AreEqual("3", cf2.Formula2);
                    Assert.AreEqual(ExcelFillStyle.Solid, cf2.Style.Fill.PatternType);
                    Assert.AreEqual(Color.Red.ToArgb(), cf2.Style.Fill.BackgroundColor.Color.Value.ToArgb());
                    SaveWorkbook("cfcopy.xlsx", p2);
                }
            }
        }

        [TestMethod]
        public void CopyComments()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);
                ws.Cells["A1"].AddComment("Comment");
                ws.Cells["A1:A2"].Copy(ws.Cells["B5:B6"], ExcelRangeCopyOptionFlags.ExcludeValues, ExcelRangeCopyOptionFlags.ExcludeStyles);

                Assert.AreEqual("Comment", ws.Cells["B5"].Comment.Text);
                Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["B6"].Formula));
                Assert.IsFalse(ws.Cells["B5"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Italic);

                ws.Cells["A1:A2"].Copy(ws.Cells["C5:C6"], ExcelRangeCopyOptionFlags.ExcludeComments);

                Assert.IsNull(ws.Cells["C5"].Comment);
                Assert.IsFalse(string.IsNullOrEmpty(ws.Cells["C6"].Formula));
                Assert.IsTrue(ws.Cells["C5"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Italic);
            }
        }
        [TestMethod]
        public void CopyThreadedComments()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);
                ws.Cells["A2"].AddThreadedComment();
                ws.Cells["A2"].ThreadedComment.AddComment("1", "Threaded Comment");
                ws.Cells["A1:A2"].Copy(ws.Cells["B5:B6"], ExcelRangeCopyOptionFlags.ExcludeValues, ExcelRangeCopyOptionFlags.ExcludeStyles);

                Assert.AreEqual("Threaded Comment", ws.Cells["B6"].ThreadedComment.Comments[0].Text);
                Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["B6"].Formula));
                Assert.IsFalse(ws.Cells["B5"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Italic);

                ws.Cells["A1:A2"].Copy(ws.Cells["C5:C6"], ExcelRangeCopyOptionFlags.ExcludeThreadedComments);

                Assert.IsNull(ws.Cells["C6"].ThreadedComment);
                Assert.IsFalse(string.IsNullOrEmpty(ws.Cells["C6"].Formula));
                Assert.IsTrue(ws.Cells["C5"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Italic);
            }
        }

        [TestMethod]
        public void CopyMergedCells()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);
                ws.Cells["A1"].AddComment("Comment");
                ws.Cells["A1:A2"].Merge = true;

                ws.Cells["A1:A2"].Copy(ws.Cells["B5:B6"], ExcelRangeCopyOptionFlags.ExcludeValues, ExcelRangeCopyOptionFlags.ExcludeStyles);

                Assert.IsTrue(ws.Cells["B5:B6"].Merge);
                Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["B6"].Formula));
                Assert.IsFalse(ws.Cells["B5"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Italic);

                ws.Cells["A1:A2"].Copy(ws.Cells["C5:C6"], ExcelRangeCopyOptionFlags.ExcludeMergedCells);

                Assert.IsFalse(ws.Cells["C5:C6"].Merge);
                Assert.IsFalse(string.IsNullOrEmpty(ws.Cells["C6"].Formula));
                Assert.IsTrue(ws.Cells["C5"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Italic);

            }
        }
        [TestMethod]
        public void CopyHyperLinks()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);
                ws.Cells["A1"].AddComment("Comment");
                ws.Cells["A2"].SetHyperlink(ws.Cells["C3"], "Link to C3");

                ws.Cells["A1:A2"].Copy(ws.Cells["B5:B6"], ExcelRangeCopyOptionFlags.ExcludeValues, ExcelRangeCopyOptionFlags.ExcludeStyles);

                Assert.AreEqual("Link to C3", ((ExcelHyperLink)ws.Cells["B6"].Hyperlink).Display);
                Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["B6"].Formula));
                Assert.IsFalse(ws.Cells["B5"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["B6"].Style.Font.Italic);

                ws.Cells["A1:A2"].Copy(ws.Cells["C5:C6"], ExcelRangeCopyOptionFlags.ExcludeHyperLinks);

                Assert.IsNull(ws.Cells["C6"].Hyperlink);
                Assert.IsFalse(string.IsNullOrEmpty(ws.Cells["C6"].Formula));
                Assert.IsTrue(ws.Cells["C5"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Italic);

            }
        }
        [TestMethod]
        public void CopyStylesWithinWorkbook()
        {
            using (var p = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p);

                string nf = "#,##0";
                ws.Cells["B1"].Style.Font.UnderLineType = ExcelUnderLineType.Double;
                ws.Cells["B2"].Style.Numberformat.Format = nf;
                ws.Cells["A1:B2"].CopyStyles(ws.Cells["C5:F8"]);

                Assert.IsTrue(ws.Cells["C5"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Bold);
                Assert.IsTrue(ws.Cells["C8"].Style.Font.Bold);
                Assert.IsFalse(ws.Cells["C9"].Style.Font.Bold);

                Assert.IsFalse(ws.Cells["C5"].Style.Font.Italic);
                Assert.IsTrue(ws.Cells["C6"].Style.Font.Italic);
                Assert.IsTrue(ws.Cells["C8"].Style.Font.Italic);
                Assert.IsFalse(ws.Cells["C9"].Style.Font.Bold);

                Assert.AreEqual(nf, ws.Cells["D6"].Style.Numberformat.Format);
                Assert.AreEqual(nf, ws.Cells["F6"].Style.Numberformat.Format);
                Assert.AreEqual(nf, ws.Cells["E8"].Style.Numberformat.Format);
                Assert.AreEqual(ExcelUnderLineType.Double, ws.Cells["D5"].Style.Font.UnderLineType);
                Assert.AreEqual(ExcelUnderLineType.Double, ws.Cells["F5"].Style.Font.UnderLineType);
                Assert.AreEqual(ExcelUnderLineType.None, ws.Cells["D6"].Style.Font.UnderLineType);
                Assert.AreEqual(ExcelUnderLineType.None, ws.Cells["F8"].Style.Font.UnderLineType);

                //SaveWorkbook("styleCopy.xlsx", p);
            }
        }
        [TestMethod]
        public void CopyStylesToNewWorkbook()
        {
            using (var p1 = new ExcelPackage())
            {
                ExcelWorksheet ws = SetupCopyRange(p1);
                using (var p2 = new ExcelPackage())
                {
                    var ws2 = p2.Workbook.Worksheets.Add("Sheet1");
                    string nf = "#,##0";
                    ws.Cells["B1"].Style.Font.UnderLineType = ExcelUnderLineType.Double;
                    ws.Cells["B2"].Style.Numberformat.Format = nf;
                    ws.Cells["A1:B2"].CopyStyles(ws2.Cells["C5:F8"]);

                    Assert.IsTrue(ws2.Cells["C5"].Style.Font.Bold);
                    Assert.IsTrue(ws2.Cells["C6"].Style.Font.Bold);
                    Assert.IsTrue(ws2.Cells["C8"].Style.Font.Bold);
                    Assert.IsFalse(ws2.Cells["C9"].Style.Font.Bold);

                    Assert.IsFalse(ws2.Cells["C5"].Style.Font.Italic);
                    Assert.IsTrue(ws2.Cells["C6"].Style.Font.Italic);
                    Assert.IsTrue(ws2.Cells["C8"].Style.Font.Italic);
                    Assert.IsFalse(ws2.Cells["C9"].Style.Font.Bold);

                    Assert.AreEqual(nf, ws2.Cells["D6"].Style.Numberformat.Format);
                    Assert.AreEqual(nf, ws2.Cells["F6"].Style.Numberformat.Format);
                    Assert.AreEqual(nf, ws2.Cells["E8"].Style.Numberformat.Format);
                    Assert.AreEqual(ExcelUnderLineType.Double, ws2.Cells["D5"].Style.Font.UnderLineType);
                    Assert.AreEqual(ExcelUnderLineType.Double, ws2.Cells["F5"].Style.Font.UnderLineType);
                    Assert.AreEqual(ExcelUnderLineType.None, ws2.Cells["D6"].Style.Font.UnderLineType);
                    Assert.AreEqual(ExcelUnderLineType.None, ws2.Cells["F8"].Style.Font.UnderLineType);

                    SaveWorkbook("styleCopy.xlsx", p2);
                }
            }
        }

        private static ExcelWorksheet SetupCopyRange(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = 1;
            ws.Cells["A2"].Formula = "A1+1";
            ws.Cells["A1:A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.Font.Italic = true;
            ws.Calculate();
            return ws;
        }
    }
}
