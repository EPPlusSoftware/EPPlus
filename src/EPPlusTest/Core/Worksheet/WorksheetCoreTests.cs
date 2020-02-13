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

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class WorksheetCoreTests : TestBase
    {
        [TestMethod]
        public void ValidateAutoFitDontShowHiddenColumns()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("AutoFitHidden");
                LoadTestdata(ws);

                ws.Column(2).Hidden = true;

                ws.Cells.AutoFitColumns(); 
                Assert.AreEqual(true, ws.Column(2).Hidden);
                p.Save();
            }
        }
        [TestMethod]
        public void RichTextFlagShouldBeCleanedWhenOverwritingValue()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RichTextOverwriteValue");

                ws.Cells["A1:B2"].RichText.Add("RichText");

                ws.Cells["A1"].Value = "Text";
                ws.Cells["B2"].Value = "Text";
                Assert.IsFalse(ws.Cells["A1"].IsRichText);
                Assert.IsTrue(ws.Cells["A2"].IsRichText);
                Assert.IsTrue(ws.Cells["B1"].IsRichText);
                Assert.IsFalse(ws.Cells["B2"].IsRichText);
            }
        }

        [TestMethod]
        public void RichTextFlagShouldBeCleanedWhenOverwritingValueAddress()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RichTextOverwriteAddress");

                ws.Cells["A1:B2"].RichText.Add("RichText");
                Assert.IsTrue(ws.Cells["A1"].IsRichText);
                Assert.IsTrue(ws.Cells["B2"].IsRichText);

                ws.Cells["A1:B2"].Value = "Text";
                Assert.IsFalse(ws.Cells["A1"].IsRichText);
                Assert.IsFalse(ws.Cells["B2"].IsRichText);
            }
        }

        [TestMethod]
        public void RichTextFlagShouldBeCleanedWhenOverwritingWithArray()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RichTextOverwrite");

                ws.Cells["A1:C3"].RichText.Add("RichText");
                Assert.IsTrue(ws.Cells["A1"].IsRichText);
                Assert.IsTrue(ws.Cells["B2"].IsRichText);

                ws.Cells["A1:B2"].Value = new string[,] { { "Text", "Text" }, { "Text", "Text" } };
                Assert.IsFalse(ws.Cells["A1"].IsRichText);
                Assert.IsFalse(ws.Cells["B2"].IsRichText);
                Assert.IsTrue(ws.Cells["C1"].IsRichText);
                Assert.IsTrue(ws.Cells["C3"].IsRichText);
                Assert.IsTrue(ws.Cells["A3"].IsRichText);
            }
        }
        [TestMethod]
        public void FormulaShouldBeCleanedWhenOverwritingWithArray()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RichTextOverwrite");

                ws.Cells["A1:C3"].FormulaR1C1 = "RC";
                Assert.IsFalse(ws.Cells["A1"].Formula==null);

                ws.Cells["A1:B2"].Value = new string[,] { { "Text", "Text" }, { "Text", "Text" } };
                Assert.IsTrue(ws.Cells["A1"].Formula == "");
            }
        }
    }
}
