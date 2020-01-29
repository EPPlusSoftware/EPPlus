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
    }
}
