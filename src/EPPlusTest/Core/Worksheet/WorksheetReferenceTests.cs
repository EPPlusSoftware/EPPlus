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
    public class WorksheetReferenceTests : TestBase
    {
        [TestMethod]
        public void InsertRowsUpdatesReferencesCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells[2, 2].Formula = "C3";
                sheet1.Cells[3, 3].Value = "Hello, world!";
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
                sheet1.InsertRow(3, 10);
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet1.Cells[13, 3].Value);
                Assert.AreEqual("C13", sheet1.Cells[2, 2].Formula);
                Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
            }
        }

        [TestMethod]
        public void CrossSheetInsertRowsUpdatesReferencesCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells[2, 2].Formula = "Sheet2!C3";
                sheet2.Cells[3, 3].Value = "Hello, world!";
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
                sheet2.InsertRow(3, 10);
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet2.Cells[13, 3].Value);
                Assert.AreEqual("Sheet2!C13", sheet1.Cells[2, 2].Formula, true);
                Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
            }
        }

        [TestMethod]
        public void CrossSheetInsertColumnsUpdatesReferencesCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells[2, 2].Formula = "'Sheet2'!C3";
                sheet2.Cells[3, 3].Value = "Hello, world!";
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
                sheet2.InsertColumn(3, 10);
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet2.Cells[3, 13].Value);
                Assert.AreEqual("'Sheet2'!M3", sheet1.Cells[2, 2].Formula);
                Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
            }
        }
        [TestMethod]
        public void CrossSheetInsertRowAfterReferencesHasNoEffect()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("New Sheet");
                var otherSheet = package.Workbook.Worksheets.Add("Other Sheet");
                sheet.Cells[3, 3].Formula = "'Other Sheet'!C3";
                otherSheet.Cells[3, 3].Formula = "45";
                otherSheet.InsertRow(5, 1);
                Assert.AreEqual("'Other Sheet'!C3", sheet.Cells[3, 3].Formula);
            }
        }

        [TestMethod]
        public void CrossSheetInsertColumnAfterReferencesHasNoEffect()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("New Sheet");
                var otherSheet = package.Workbook.Worksheets.Add("Other Sheet");
                sheet.Cells[3, 3].Formula = "'Other Sheet'!C3";
                otherSheet.Cells[3, 3].Formula = "45";
                otherSheet.InsertColumn(5, 1);
                Assert.AreEqual("'Other Sheet'!C3", sheet.Cells[3, 3].Formula);
            }
        }

        [TestMethod]
        public void CrossSheetReferenceIsUpdatedWhenSheetIsRenamed()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("New Sheet");
                var otherSheet = package.Workbook.Worksheets.Add("Other Sheet");
                sheet.Cells[3, 3].Formula = "'Other Sheet'!C3";
                otherSheet.Cells[3, 3].Formula = "45";
                otherSheet.Name = "New Name";
                Assert.AreEqual("'New Name'!C3", sheet.Cells[3, 3].Formula);
            }
        }

        [TestMethod]
        public void CopyCellUpdatesRelativeCrossSheetReferencesCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells[3, 3].Formula = "Sheet2!C3";
                sheet2.Cells[3, 3].Value = "Hello, world!";
                sheet2.Cells[3, 4].Value = "Hello, WORLD!";
                sheet2.Cells[4, 3].Value = "Goodbye, world!";
                sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
                sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
                package.Workbook.Calculate();
                Assert.AreEqual("Sheet2!D4", sheet1.Cells[4, 4].Formula, true);
                Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
                Assert.AreEqual("Goodbye, WORLD!", sheet1.Cells[4, 4].Value);
            }
        }

        [TestMethod]
        public void CopyCellUpdatesAbsoluteCrossSheetReferencesCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells[3, 3].Formula = "Sheet2!$C$3";
                sheet2.Cells[3, 3].Value = "Hello, world!";
                sheet2.Cells[3, 4].Value = "Hello, WORLD!";
                sheet2.Cells[4, 3].Value = "Goodbye, world!";
                sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
                sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
                package.Workbook.Calculate();
                Assert.AreEqual("Sheet2!$C$3", sheet1.Cells[4, 4].Formula, true);
                Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
                Assert.AreEqual("Hello, world!", sheet1.Cells[4, 4].Value);
            }
        }

        [TestMethod]
        public void CopyCellUpdatesRowAbsoluteCrossSheetReferencesCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells[3, 3].Formula = "Sheet2!C$3";
                sheet2.Cells[3, 3].Value = "Hello, world!";
                sheet2.Cells[3, 4].Value = "Hello, WORLD!";
                sheet2.Cells[4, 3].Value = "Goodbye, world!";
                sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
                sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
                package.Workbook.Calculate();
                Assert.AreEqual("Sheet2!D$3", sheet1.Cells[4, 4].Formula, true);
                Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
                Assert.AreEqual("Hello, WORLD!", sheet1.Cells[4, 4].Value);
            }
        }

        [TestMethod]
        public void CopyCellUpdatesColumnAbsoluteCrossSheetReferencesCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells[3, 3].Formula = "Sheet2!$C3";
                sheet2.Cells[3, 3].Value = "Hello, world!";
                sheet2.Cells[3, 4].Value = "Hello, WORLD!";
                sheet2.Cells[4, 3].Value = "Goodbye, world!";
                sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
                package.Workbook.Calculate();
                Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
                sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
                package.Workbook.Calculate();
                Assert.AreEqual("Sheet2!$C4", sheet1.Cells[4, 4].Formula, true);
                Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
                Assert.AreEqual("Goodbye, world!", sheet1.Cells[4, 4].Value);
            }
        }
        [TestMethod]
        public void AddWorksheetWithDollarSign()
        {
            using (var package = new ExcelPackage())
            {
                var sheetName = "Sheet1$";
                var sheet1 = package.Workbook.Worksheets.Add(sheetName);
                Assert.AreEqual(sheetName, sheet1.Name);
                Assert.AreEqual(sheet1, package.Workbook.Worksheets[sheetName]);

                package.Workbook.Names.Add("WorkbookDefinedName1", sheet1.Cells["A1"], true);
                sheet1.Names.Add("DefinedName1", sheet1.Cells["A1"], true);

                package.Save();
                using (var p2 = new ExcelPackage(package.Stream))
                {
                    sheet1 = package.Workbook.Worksheets[sheetName];
                    Assert.IsNotNull(sheet1);
                    Assert.AreEqual(sheetName, sheet1.Name);
                    Assert.AreEqual(sheet1, package.Workbook.Worksheets[sheetName]);
                    Assert.AreEqual("'Sheet1$'!A1", sheet1.Names["DefinedName1"].FullAddress);
                }
            }
        }
        [TestMethod]
        public void AddDefinedNameRanges()
        {
            using (var package = new ExcelPackage())
            {
                var sheetName = "AName";
                var sheet1 = package.Workbook.Worksheets.Add(sheetName);
                Assert.AreEqual(sheetName, sheet1.Name);
                Assert.AreEqual(sheet1, package.Workbook.Worksheets[sheetName]);

                package.Workbook.Names.Add("NamedRange1", sheet1.Cells["A1:C10"], false);
                sheet1.Names.Add("NamedRange2", sheet1.Cells["A1:Z10"], false);

                package.Save();

                using (var p2 = new ExcelPackage(package.Stream))
                {
                    sheet1 = package.Workbook.Worksheets[sheetName];
                    Assert.IsNotNull(sheet1);
                    Assert.AreEqual(sheetName, sheet1.Name);
                    Assert.AreEqual(sheet1, package.Workbook.Worksheets[sheetName]);
                    Assert.AreEqual("AName!$A$1:$Z$10", sheet1.Names["NamedRange2"].FullAddress);
                    Assert.AreEqual("AName!$A$1:$C$10", p2.Workbook.Names["NamedRange1"].FullAddress);
                }
            }
        }
        [TestMethod]
        public void VerifyAddressFullAddress()
        {
            using (var package = new ExcelPackage())
            {
                //var ws1 = package.Workbook.Worksheets.Add("sheet1");
                //var n = ws1.Names.Add("Name1", ws1.Cells["A1:B5"]);
                //n.Address = "A2:D6";
                //Assert.AreEqual("sheet1!A2:D6", n.FullAddress);
                //Assert.AreEqual("sheet1!$A$2:$D$6", n.FullAddressAbsolute);
                var ws1 = package.Workbook.Worksheets.Add("sheet1");
                var n = package.Workbook.Names.Add("Name1", ws1.Cells["A1:B5"]);
                n.Address = "A2:D6";
                Assert.AreEqual("sheet1!A2:D6", n.FullAddress);
                Assert.AreEqual("sheet1!$A$2:$D$6", n.FullAddressAbsolute);


            }
        }

        [TestMethod]
        public void MoveToStart_IsWorksheets1BasedFalse_NoException()
        {
            using var p = new ExcelPackage();
            p.Compatibility.IsWorksheets1Based = false;

            var ws1 = p.Workbook.Worksheets.Add("Sheet 1");
            var ws2 = p.Workbook.Worksheets.Add("Sheet 2");

            p.Workbook.Worksheets.MoveToStart(ws2.Name);
        }

        [TestMethod]
        public void MoveToStart_IsWorksheets1BasedTrue_Exception()
        {
            using var p = new ExcelPackage();
            p.Compatibility.IsWorksheets1Based = true;

            var ws1 = p.Workbook.Worksheets.Add("Sheet 1");
            var ws2 = p.Workbook.Worksheets.Add("Sheet 2");

            p.Workbook.Worksheets.MoveToStart(ws2.Name);
        }
    }
}
