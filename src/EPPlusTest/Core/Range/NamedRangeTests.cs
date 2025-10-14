using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System;
using System.IO;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class NamedRangeTests : TestBase
    {
        [TestMethod]
        public void IsValidName()
        {
            Assert.IsFalse(ExcelAddressUtil.IsValidName("123sa"));  //invalid start char 
            Assert.IsFalse(ExcelAddressUtil.IsValidName("*d"));     //invalid start char
            Assert.IsFalse(ExcelAddressUtil.IsValidName("\t"));     //invalid start char
            Assert.IsFalse(ExcelAddressUtil.IsValidName("\\t"));    //Backslash at least three chars
            Assert.IsFalse(ExcelAddressUtil.IsValidName("A+1"));   //invalid char
            Assert.IsFalse(ExcelAddressUtil.IsValidName("A%we"));   //Address invalid
            Assert.IsFalse(ExcelAddressUtil.IsValidName("BB73"));   //Address invalid
            Assert.IsTrue(ExcelAddressUtil.IsValidName("\\tr"));    //Backslash at least three chars
            Assert.IsTrue(ExcelAddressUtil.IsValidName("BBBB75"));  //Valid
            Assert.IsTrue(ExcelAddressUtil.IsValidName("BB1500005")); //Valid
        }
        [TestMethod]
        public void NamedRangeMovesDownIfRowInsertedAbove()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 1, 3, 3];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertRow(1, 1);

                Assert.AreEqual("NEW!$A$3:$C$4", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeDoesNotChangeIfRowInsertedBelow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 1, 3, 3];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertRow(4, 1);

                Assert.AreEqual("$A$2:$C$3", namedRange.Address);
            }
        }

        [TestMethod]
        public void NamedRangeExpandsDownIfRowInsertedWithin()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 1, 3, 3];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertRow(3, 1);

                Assert.AreEqual("NEW!$A$2:$C$4", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeMovesRightIfColInsertedBefore()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 2, 3, 4];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertColumn(1, 1);

                Assert.AreEqual("NEW!$C$2:$E$3", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeUnchangedIfColInsertedAfter()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 2, 3, 4];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertColumn(5, 1);

                Assert.AreEqual("$B$2:$D$3", namedRange.Address);
            }
        }

        [TestMethod]
        public void NamedRangeExpandsToRightIfColInsertedWithin()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 2, 3, 4];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertColumn(5, 1);

                Assert.AreEqual("$B$2:$D$3", namedRange.Address);
            }
        }

        [TestMethod]
        public void NamedRangeWithWorkbookScopeIsMovedDownIfRowInsertedAbove()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 1, 3, 3];
                var namedRange = workbook.Names.Add("NewNamedRange", range);

                sheet.InsertRow(1, 1);

                Assert.AreEqual("NEW!$A$3:$C$4", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeWithWorkbookScopeIsMovedRightIfColInsertedBefore()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 2, 3, 3];
                var namedRange = workbook.Names.Add("NewNamedRange", range);

                sheet.InsertColumn(1, 1);

                Assert.AreEqual("NEW!$C$2:$D$3", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeIsUnchangedForOutOfScopeSheet()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet1 = package.Workbook.Worksheets.Add("NEW");
                var sheet2 = package.Workbook.Worksheets.Add("NEW2");
                var range = sheet2.Cells[2, 2, 3, 3];
                var namedRange = workbook.Names.Add("NewNamedRange", range);

                sheet1.InsertColumn(1, 1);

                Assert.AreEqual("$B$2:$C$3", namedRange.Address);
            }
        }
        [TestMethod]
        public void NamedRangeIsEqual()
        {
            using (var p1 = new ExcelPackage())
            {
                using (var p2 = new ExcelPackage())
                {
                    var ws1 = p1.Workbook.Worksheets.Add("sheet1");
                    var ws2 = p1.Workbook.Worksheets.Add("sheet2");

                    var ws1_p2 = p2.Workbook.Worksheets.Add("sheet1");


                    var wbName1 = p1.Workbook.Names.Add("Name1", ws1.Cells["sheet1!A1"]);
                    var wsName1 = ws1.Names.Add("Name1", ws1.Cells["A1"]);
                    var wsName2 = ws1.Names.Add("Name2", ws1.Cells["A1"]);

                    var wsName1_p2 = ws1_p2.Names.Add("Name1", ws1_p2.Cells["A1"]);

                    //Assert
                    Assert.IsTrue(wbName1.Equals(wbName1));
                    Assert.IsTrue(wsName1.Equals(wsName1));

                    Assert.IsFalse(wsName1.Equals(wbName1));
                    Assert.IsFalse(wbName1.Equals(wsName2));
                    Assert.IsFalse(wsName1.Equals(wsName1_p2));
                }
            }
        }

        [TestMethod]
        public void WorkbookNamedRange_ShouldRetain_FixedAddress()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    package.Workbook.Names.Add("MyName", sheet.Cells["$A$1:$A$3"]);
                    package.Save();
                }
                ms.Position = 0;
                using (var package2 = new ExcelPackage(ms))
                {
                    var nameAddress = package2.Workbook.Names["MyName"].ToInternalAddress().Address;
                    Assert.AreEqual("test!$A$1:$A$3", nameAddress);
                }
            }
        }

        [TestMethod]
        public void WorksheetNamedRange_ShouldRetain_FixedAddress()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    sheet.Names.Add("MyName", sheet.Cells["$A$1:$A$3"]);
                    package.Save();
                }
                ms.Position = 0;
                using (var package2 = new ExcelPackage(ms))
                {
                    var nameAddress = package2.Workbook.Worksheets["test"].Names["MyName"].ToInternalAddress().Address;
                    Assert.AreEqual("test!$A$1:$A$3", nameAddress);
                }
            }
        }

        [TestMethod]
        public void WorkbookNamedRange_ShouldRetainRelativeAddress_WhenIsRelativeIsTrue()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    var n = package.Workbook.Names.Add("MyName", sheet.Cells["A1:A3"], true);
                    package.Save();
                }
                ms.Position = 0;
                using (var package2 = new ExcelPackage(ms))
                {
                    var nameAddress = package2.Workbook.Names["MyName"].ToInternalAddress().Address;
                    Assert.AreEqual("test!A1:A3", nameAddress);
                }
            }
        }

        [TestMethod]
        public void WorksheetNamedRange_ShouldRetainRelativeAddress_WhenIsRelativeIsTrue()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    sheet.Names.Add("MyName", sheet.Cells["A1:A3"], true);
                    package.Save();
                }
                ms.Position = 0;
                using (var package2 = new ExcelPackage(ms))
                {
                    var nameAddress = package2.Workbook.Worksheets["test"].Names["MyName"].ToInternalAddress().Address;
                    Assert.AreEqual("test!A1:A3", nameAddress);
                }
            }
        }

        [TestMethod]
        public void WorkbookNamedRange_ShouldNotRetainRelativeAddress_WhenIsRelativeIsFalse()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    package.Workbook.Names.Add("MyName", sheet.Cells["A1:A3"], false);
                    package.Save();
                }
                ms.Position = 0;
                using (var package2 = new ExcelPackage(ms))
                {
                    var nameAddress = package2.Workbook.Names["MyName"].ToInternalAddress().Address;
                    Assert.AreEqual("test!$A$1:$A$3", nameAddress);
                }
            }
        }

        [TestMethod]
        public void WorksheetNamedRange_ShouldNotRetainRelativeAddress_WhenIsRelativeIsFalse()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    sheet.Names.Add("MyName", sheet.Cells["A1:A3"], false);
                    package.Save();
                }
                ms.Position = 0;
                using (var package2 = new ExcelPackage(ms))
                {
                    var nameAddress = package2.Workbook.Worksheets["test"].Names["MyName"].ToInternalAddress().Address;
                    Assert.AreEqual("test!$A$1:$A$3", nameAddress);
                }
            }
        }

        [TestMethod]
        public void WorkbookNamedRange_ShouldAlwaysSetFixedAddress_WhenNotLoadingFromFile()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    package.Workbook.Names.Add("MyName", sheet.Cells["A1:A3"]);
                    package.Save();
                }
                ms.Position = 0;
                using (var package2 = new ExcelPackage(ms))
                {
                    var nameAddress = package2.Workbook.Names["MyName"].ToInternalAddress().Address;
                    Assert.AreEqual("test!$A$1:$A$3", nameAddress);
                }
            }
        }

        [TestMethod]
        public void WorksheetNamedRange_ShouldAlwaysSetFixedAddress_WhenNotLoadingFromFile()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    sheet.Names.Add("MyName", sheet.Cells["A1:A3"]);
                    package.Save();
                }
                ms.Position = 0;
                using (var package2 = new ExcelPackage(ms))
                {
                    var nameAddress = package2.Workbook.Worksheets["test"].Names["MyName"].ToInternalAddress().Address;
                    Assert.AreEqual("test!$A$1:$A$3", nameAddress);
                }
            }
        }
        [TestMethod]
        public void CopyWorksheetWithNamePointingToAnotherSheet()
        {
            using (var pck = new ExcelPackage())
            {
                // Add two worksheets
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = pck.Workbook.Worksheets.Add("Sheet2");

                // Add a name scoped to sheet 1 that points to sheet 2
                sheet1.Names.Add("Name1", sheet2.Cells["A1"]);

                // Create a new workbook
                using (var newPck = new ExcelPackage())
                {
                    // Copy sheet 1 to the new workbook
                    newPck.Workbook.Worksheets.Add("Sheet1", sheet1);
                    var copiedSheet1 = newPck.Workbook.Worksheets["Sheet1"];
                    Assert.IsNotNull(copiedSheet1);
                    Assert.AreEqual(1, copiedSheet1.Names.Count);
                    Assert.AreEqual("#REF!", copiedSheet1.Names[0].NameFormula);

                    newPck.Save();

                    using (var newPck2 = new ExcelPackage(newPck.Stream))
                    {
                        var wsSaved = newPck2.Workbook.Worksheets[0];
                        Assert.AreEqual(1, wsSaved.Names.Count);
                        Assert.AreEqual("#REF!", wsSaved.Names[0].NameFormula);
                    }
                }
            }
        }
        [TestMethod]
        public void AddNamesShouldPassValidationTest()
        {
            using (var pck = new ExcelPackage())
            {
                // Add two worksheets
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = pck.Workbook.Worksheets.Add("Sheet 2");
                //p.Workbook.ExternalLinks.AddExternalWorkbook(new FileInfo("c:\\temp\\arc.xlsx"));
                var wb = pck.Workbook;
                //Workbook level.
                wb.Names.AddFormula("NameWithFormula1", "Sheet1!A1 * 'Sheet 2'!A1");
                wb.Names.AddFormula("NameWithFormula2", "Sheet1!A1 * ('Sheet 2'!A1 + Sheet3!A4) / 'Sheet 2'!A8"); //Missing sheet3 should pass
                wb.Names.AddFormula("NameWithFormula3", "Sheet1!A1 * ('Sheet 2'!A1 + Sheet1!Name1)");
                wb.Names.AddFormula("NameWithFormula4", "([0]ExternalSheet1!A1 * ('Sheet 2'!A1 + Name1))+1"); //External reference.
                wb.Names.AddFormula("NameWithFormula5", "(Sum([0]ExternalSheet1!A1:A8) - (Avg('Sheet 2'!A1:B12) + NameWithFormula1))+1"); //External reference.
                wb.Names.AddFormula("NameWithFormula6", "Sum(#REF!) - Avg('Sheet 2'!#REF!)"); //External reference.
                SaveWorkbook("NamesShouldpass.xlsx", pck);
            }
        }
        [TestMethod]
        public void CopyNameToNewWorkbookFromRange()
        {
            using (var p = new ExcelPackage())
            {
                // Add two worksheets
                var sheet1 = p.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = p.Workbook.Worksheets.Add("Sheet 2");
                //p.Workbook.ExternalLinks.AddExternalWorkbook(new FileInfo("c:\\temp\\arc.xlsx"));
                var wb = p.Workbook;
                //Workbook level.
                wb.Names.AddValue("NameWithFormula1", 1);
                wb.Names.Add("NameWithFormulaAddress", sheet1.Cells["A5"]);
                sheet1.Names.AddValue("NameWithFormulaWs1", 2);
                sheet1.Cells["A2"].Formula = "NameWithFormula1+1";
                sheet1.Cells["A3"].Formula = "NameWithFormulaAddress+NameWithFormula1";
                sheet1.Cells["A4"].Formula = "NameWithFormulaAddress+NameWithFormula1-NameWithFormulaWs1";
                using (var p2 = new ExcelPackage())
                {
                    var sheetNew = p2.Workbook.Worksheets.Add("Sheet1");
                    sheet1.Cells["A1:A4"].Copy(sheetNew.Cells["A1"]);
                    Assert.AreEqual(2, p2.Workbook.Names.Count);
                    Assert.AreEqual(1, sheetNew.Names.Count);
                    SaveWorkbook("CopyWithName.xlsx", p2);
                }
            }
        }
        [TestMethod]
        public void CopyNameToNewWorkbook()
        {
            using (var p = new ExcelPackage())
            {
                // Add two worksheets
                var sheet1 = p.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = p.Workbook.Worksheets.Add("Sheet 2");
                var wb = p.Workbook;
                //Workbook level.
                wb.Names.AddValue("NameWithFormula1", 1);
                wb.Names.Add("NameWithFormulaAddress", sheet1.Cells["A5"]);
                sheet1.Names.AddValue("NameWithFormulaWs1", 2);
                sheet1.Cells["A2"].Formula = "NameWithFormula1+1";
                sheet1.Cells["A3"].Formula = "NameWithFormulaAddress+NameWithFormula1";
                sheet1.Cells["A4"].Formula = "NameWithFormulaAddress+NameWithFormula1-NameWithFormulaWs1";
                using (var p2 = new ExcelPackage())
                {
                    var sheetNew=p2.Workbook.Worksheets.Add("Sheet1", sheet1);
                    Assert.AreEqual(2, p2.Workbook.Names.Count);
                    Assert.AreEqual(1, sheetNew.Names.Count);
                    SaveWorkbook("CopyWithName.xlsx", p2);
                }
            }
        }
        [TestMethod]
        public void AddRowToDefinedNamesWithMultipleAddresses()
        {
            using (var p = new ExcelPackage())
            {
                // Add two worksheets
                var sheet1 = p.Workbook.Worksheets.Add("Sheet1");
                var wb = p.Workbook;
                //Workbook level.
                wb.Names.Add("WbName", sheet1.Cells["A5:F6,C7:D11,D4"]);
                sheet1.Names.Add("SheetName", sheet1.Cells["A5,D7:D11,D1:D7"]);
                sheet1.InsertRow(6, 2);

                Assert.AreEqual("$A$5:$F$8,$C$9:$D$13,$D$4", wb.Names["wbName"].Address);
                Assert.AreEqual("$A$5,$D$9:$D$13,$D$1:$D$9", sheet1.Names["sheetName"].Address);
            }
        }
        [TestMethod]
        public void DeleteRowToDefinedNamesWithMultipleAddresses()
        {
            using (var p = new ExcelPackage())
            {
                // Add two worksheets
                var sheet1 = p.Workbook.Worksheets.Add("Sheet1");
                var wb = p.Workbook;
                //Workbook level.
                wb.Names.Add("WbName", sheet1.Cells["A5:F6,C7:D11,D4"]);
                sheet1.Names.Add("SheetName", sheet1.Cells["A5,D7:D11,D1:D7"]);
                sheet1.DeleteRow(6, 2);

                Assert.AreEqual("$A$5:$F$5,$C$6:$D$9,$D$4", wb.Names["wbName"].Address);
                Assert.AreEqual("$A$5,$D$6:$D$9,$D$1:$D$5", sheet1.Names["sheetName"].Address);
            }
        }
        [TestMethod]
        public void DeleteColumnToDefinedNamesWithMultipleAddresses()
        {
            using (var p = new ExcelPackage())
            {
                // Add two worksheets
                var sheet1 = p.Workbook.Worksheets.Add("Sheet1");
                var wb = p.Workbook;
                //Workbook level.
                wb.Names.Add("WbName", sheet1.Cells["A5:F6,C7:D11,D4"]);
                sheet1.Names.Add("SheetName", sheet1.Cells["A5,D7:D11,D1:D7"]);
                sheet1.DeleteColumn(2, 2);

                Assert.AreEqual("$A$5:$D$6,$B$7:$B$11,$B$4", wb.Names["wbName"].Address);
                Assert.AreEqual("$A$5,$B$7:$B$11,$B$1:$B$7", sheet1.Names["sheetName"].Address);
            }
        }
        [TestMethod]
        public void AddColumnToDefinedNamesWithMultipleAddresses()
        {
            using (var p = new ExcelPackage())
            {
                // Add two worksheets
                var sheet1 = p.Workbook.Worksheets.Add("Sheet1");
                var wb = p.Workbook;
                //Workbook level.
                wb.Names.Add("WbName", sheet1.Cells["A5:F6,C7:D11,D4"]);
                sheet1.Names.Add("SheetName", sheet1.Cells["A5,D7:D11,D1:D7"]);
                sheet1.InsertColumn(2, 2);

                Assert.AreEqual("$A$5:$H$6,$E$7:$F$11,$F$4", wb.Names["wbName"].Address);
                Assert.AreEqual("$A$5,$F$7:$F$11,$F$1:$F$7", sheet1.Names["sheetName"].Address);
            }
        }
    }
}
