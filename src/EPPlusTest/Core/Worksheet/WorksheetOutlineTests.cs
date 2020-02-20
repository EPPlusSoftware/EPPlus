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
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class WorksheetOutlineTests
    {
        [TestMethod]
        public void InsertRowsSetsOutlineLevel()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                sheet1.Row(15).OutlineLevel = 1;
                sheet1.InsertRow(2, 10, 15);
                for (int i = 2; i < 12; i++)
                {
                    Assert.AreEqual(1, sheet1.Row(i).OutlineLevel, $"The outline level of row {i} is not set.");
                }
                Assert.AreEqual(1, sheet1.Row(25).OutlineLevel);
            }
        }

        [TestMethod]
        public void InsertColumnsSetsOutlineLevel()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                sheet1.Column(15).OutlineLevel = 1;
                sheet1.InsertColumn(2, 10, 15);
                for (int i = 2; i < 12; i++)
                {
                    Assert.AreEqual(1, sheet1.Column(i).OutlineLevel, $"The outline level of column {i} is not set.");
                }
                Assert.AreEqual(1, sheet1.Column(25).OutlineLevel);
            }
        }

        [TestMethod]
        public void CopyRowSetsOutlineLevelsCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                sheet1.Row(2).OutlineLevel = 1;
                sheet1.Row(3).OutlineLevel = 1;
                sheet1.Row(4).OutlineLevel = 0;

                // Set outline levels on rows to be copied over.
                sheet1.Row(6).OutlineLevel = 17;
                sheet1.Row(7).OutlineLevel = 25;
                sheet1.Row(8).OutlineLevel = 29;

                sheet1.Cells["2:4"].Copy(sheet1.Cells["A6"]);
                Assert.AreEqual(1, sheet1.Row(2).OutlineLevel);
                Assert.AreEqual(1, sheet1.Row(3).OutlineLevel);
                Assert.AreEqual(0, sheet1.Row(4).OutlineLevel);

                Assert.AreEqual(1, sheet1.Row(6).OutlineLevel);
                Assert.AreEqual(1, sheet1.Row(7).OutlineLevel);
                Assert.AreEqual(0, sheet1.Row(8).OutlineLevel);
            }
        }

        [TestMethod]
        public void CopyRowCrossSheetSetsOutlineLevelsCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                sheet1.Row(2).OutlineLevel = 1;
                sheet1.Row(3).OutlineLevel = 1;
                sheet1.Row(4).OutlineLevel = 0;

                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                // Set outline levels on rows to be copied over.
                sheet2.Row(6).OutlineLevel = 17;
                sheet2.Row(7).OutlineLevel = 25;
                sheet2.Row(8).OutlineLevel = 29;

                sheet1.Cells["2:4"].Copy(sheet2.Cells["A6"]);
                Assert.AreEqual(1, sheet1.Row(2).OutlineLevel);
                Assert.AreEqual(1, sheet1.Row(3).OutlineLevel);
                Assert.AreEqual(0, sheet1.Row(4).OutlineLevel);

                Assert.AreEqual(1, sheet2.Row(6).OutlineLevel);
                Assert.AreEqual(1, sheet2.Row(7).OutlineLevel);
                Assert.AreEqual(0, sheet2.Row(8).OutlineLevel);
            }
        }

        [TestMethod]
        public void CopyColumnSetsOutlineLevelsCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                sheet1.Column(2).OutlineLevel = 1;
                sheet1.Column(3).OutlineLevel = 1;
                sheet1.Column(4).OutlineLevel = 0;

                // Set outline levels on rows to be copied over.
                sheet1.Column(6).OutlineLevel = 17;
                sheet1.Column(7).OutlineLevel = 25;
                sheet1.Column(8).OutlineLevel = 29;

                sheet1.Cells["B:D"].Copy(sheet1.Cells["F1"]);
                Assert.AreEqual(1, sheet1.Column(2).OutlineLevel);
                Assert.AreEqual(1, sheet1.Column(3).OutlineLevel);
                Assert.AreEqual(0, sheet1.Column(4).OutlineLevel);

                Assert.AreEqual(1, sheet1.Column(6).OutlineLevel);
                Assert.AreEqual(1, sheet1.Column(7).OutlineLevel);
                Assert.AreEqual(0, sheet1.Column(8).OutlineLevel);
            }
        }
    }
}
