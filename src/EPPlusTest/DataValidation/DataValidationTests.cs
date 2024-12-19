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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class DataValidationTests : ValidationTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            SetupTestData();
        }

        [TestCleanup]
        public void Cleanup()
        {
            CleanupTestData();
        }

        [TestMethod]
        public void DataValidations_ShouldNotThrowIfOperatorIsEqualAndFormula1IsEmpty()
        {
            var validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.Operator = ExcelDataValidationOperator.equal;

            validation.Validate();
        }


        [TestMethod]
        public void DataValidation_CanReadNoneValidation()
        {
            var pck = OpenTemplatePackage("i888.xlsx");

            var validation = pck.Workbook.Worksheets[0].DataValidations[0];

            Assert.IsNotNull(validation);
            Assert.AreEqual("A1", validation.Address.ToString());
            Assert.AreEqual("test", validation.PromptTitle);
            Assert.AreEqual("message", validation.Prompt);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadAllValidOperatorsOnAllTypes()
        {

        }

        public void TestTypeOperator(ExcelDataValidation type)
        {

        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void TestRangeAddMultipleTryAddingAfterShouldThrow()
        {
            ExcelPackage pck = OpenTemplatePackage("ValidationRangeTestMany.xlsx");

            var validations = pck.Workbook.Worksheets[0].DataValidations;

            validations.AddIntegerValidation("C8");
        }

        [TestMethod]
        public void TestRangeAddMultipleTryAddingAfterShouldNotThrow()
        {
            ExcelPackage pck = OpenTemplatePackage("ValidationRangeTestMany.xlsx");

            var validations = pck.Workbook.Worksheets[0].DataValidations;

            validations.AddIntegerValidation("Z8");
        }


        [TestMethod]
        public void TestRangeAddsMultipleInbetweenInstances()
        {
            ExcelPackage pck = OpenTemplatePackage("ValidationRangeTestMany.xlsx");

            var validations = pck.Workbook.Worksheets[0].DataValidations;

            StringBuilder sb = new StringBuilder();

            //Ensure all addresses exist in _validationsRD
            for (int i = 0; i < validations.Count; i++)
            {
                if (validations[i].Address.Addresses != null)
                {
                    var addresses = validations[i].Address.Addresses;

                    for (int j = 0; j < validations[i].Address.Addresses.Count; j++)
                    {
                        if (!validations._validationsRD.Exists(addresses[j]._fromRow, addresses[j]._fromCol, addresses[j]._toRow, addresses[j]._toCol))
                        {
                            sb.Append(addresses[j] + ",");
                        }
                    }
                }
                else
                {
                    if (!validations._validationsRD.Exists(validations[i].Address._fromRow, validations[i].Address._fromCol, validations[i].Address._toRow, validations[i].Address._toCol))
                    {
                        sb.Append(validations[i].Address + ",");
                    }
                }
            }

            Assert.AreEqual("", sb.ToString());
        }

        [TestMethod]
        public void TestRangeAddsSingularInstance()
        {
            ExcelPackage pck = OpenTemplatePackage("ValidationRangeTest.xlsx");

            var validations = pck.Workbook.Worksheets[0].DataValidations;

            StringBuilder sb = new StringBuilder();

            //Ensure all addresses exist in _validationsRD
            for (int i = 0; i < validations.Count; i++)
            {
                if (validations[i].Address.Addresses != null)
                {
                    var addresses = validations[i].Address.Addresses;

                    for (int j = 0; j < validations[i].Address.Addresses.Count; j++)
                    {
                        if (!validations._validationsRD.Exists(addresses[j]._fromRow, addresses[j]._fromCol, addresses[j]._toRow, addresses[j]._toCol))
                        {
                            sb.Append(addresses[i]);
                        }
                    }
                }
                else
                {
                    if (!validations._validationsRD.Exists(validations[i].Address._fromRow, validations[i].Address._fromCol, validations[i].Address._toRow, validations[i].Address._toCol))
                    {
                        sb.Append(validations[i].Address);
                    }
                }
            }

            Assert.AreEqual("", sb.ToString());
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadTypes()
        {
            var P = new ExcelPackage(new MemoryStream());
            var sheet = P.Workbook.Worksheets.Add("NewSheet");

            sheet.DataValidations.AddAnyValidation("A1");
            var intDV = sheet.DataValidations.AddIntegerValidation("A2");
            intDV.Formula.Value = 1;
            intDV.Formula2.Value = 1;

            var decimalDV = sheet.DataValidations.AddDecimalValidation("A3");

            decimalDV.Formula.Value = 1;
            decimalDV.Formula2.Value = 1;

            var listDV = sheet.DataValidations.AddListValidation("A4");

            listDV.Formula.Values.Add("5");
            listDV.Formula.Values.Add("Option");


            var textDV = sheet.DataValidations.AddTextLengthValidation("A5");

            textDV.Formula.Value = 1;
            textDV.Formula2.Value = 1;

            var dateTimeDV = sheet.DataValidations.AddDateTimeValidation("A6");

            dateTimeDV.Formula.Value = DateTime.MaxValue;
            dateTimeDV.Formula2.Value = DateTime.MinValue;

            var timeDV = sheet.DataValidations.AddTimeValidation("A7");

            timeDV.Formula.Value.Hour = 1;
            timeDV.Formula2.Value.Hour = 2;

            var customValidation = sheet.DataValidations.AddCustomValidation("A8");
            customValidation.Formula.ExcelFormula = "A1+A2";

            MemoryStream xmlStream = new MemoryStream();
            P.SaveAs(xmlStream);

            var P2 = new ExcelPackage(xmlStream);
            ExcelDataValidationCollection dataValidations = P2.Workbook.Worksheets[0].DataValidations;

            Assert.AreEqual(dataValidations[0].ValidationType.Type, eDataValidationType.Any);
            Assert.AreEqual(dataValidations[1].ValidationType.Type, eDataValidationType.Whole);
            Assert.AreEqual(dataValidations[2].ValidationType.Type, eDataValidationType.Decimal);
            Assert.AreEqual(dataValidations[3].ValidationType.Type, eDataValidationType.List);
            Assert.AreEqual(dataValidations[4].ValidationType.Type, eDataValidationType.TextLength);
            Assert.AreEqual(dataValidations[5].ValidationType.Type, eDataValidationType.DateTime);
            Assert.AreEqual(dataValidations[6].ValidationType.Type, eDataValidationType.Time);
            Assert.AreEqual(dataValidations[7].ValidationType.Type, eDataValidationType.Custom);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadOperator()
        {
            var P = new ExcelPackage(new MemoryStream());
            var sheet = P.Workbook.Worksheets.Add("NewSheet");

            var validation = sheet.DataValidations.AddIntegerValidation("A1");
            validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 1;

            MemoryStream xmlStream = new MemoryStream();
            P.SaveAs(xmlStream);

            var P2 = new ExcelPackage(xmlStream);
            Assert.AreEqual(P2.Workbook.Worksheets[0].DataValidations[0].Operator, ExcelDataValidationOperator.greaterThanOrEqual);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadShowErrorMessage()
        {
            var P = new ExcelPackage(new MemoryStream());
            var sheet = P.Workbook.Worksheets.Add("NewSheet");

            var validation = sheet.DataValidations.AddIntegerValidation("A1");

            validation.ShowErrorMessage = true;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 1;

            MemoryStream xmlStream = new MemoryStream();
            P.SaveAs(xmlStream);

            var P2 = new ExcelPackage(xmlStream);
            Assert.IsTrue(P2.Workbook.Worksheets[0].DataValidations[0].ShowErrorMessage);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadShowinputMessage()
        {
            var package = new ExcelPackage(new MemoryStream());
            CreateSheetWithIntegerValidation(package).ShowInputMessage = true;
            Assert.IsTrue(ReadIntValidation(package).ShowInputMessage);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadPrompt()
        {
            var package = new ExcelPackage(new MemoryStream());
            CreateSheetWithIntegerValidation(package).Prompt = "Prompt";
            Assert.AreEqual("Prompt", ReadIntValidation(package).Prompt);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadError()
        {
            var package = new ExcelPackage(new MemoryStream());
            var validation = CreateSheetWithIntegerValidation(package).Error = "Error";

            Assert.AreEqual("Error", ReadIntValidation(package).Error);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadErrorTitle()
        {
            var package = new ExcelPackage(new MemoryStream());
            CreateSheetWithIntegerValidation(package).ErrorTitle = "ErrorTitle";
            Assert.AreEqual("ErrorTitle", ReadIntValidation(package).ErrorTitle);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsBetweenAndFormula2IsEmpty()
        {
            var validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.Formula.Value = 1;
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Validate();
        }

        [TestMethod]
        public void DataValidations_ShouldAcceptOneItemOnly()
        {
            var validation = _sheet.DataValidations.AddListValidation("A1");
            validation.Formula.Values.Add("1");
            validation.Validate();
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfAllowBlankIsNotSet()
        {
            var validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.Validate();
        }

        [TestMethod]
        public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsOneColumn()
        {
            // Act
            var validation = _sheet.DataValidations.AddIntegerValidation("A:A");

            // Assert
            Assert.AreEqual("A1:A" + ExcelPackage.MaxRows.ToString(), validation.Address.Address);
        }

        [TestMethod]
        public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsDifferentColumns()
        {
            // Act
            var validation = _sheet.DataValidations.AddIntegerValidation("A:B");

            // Assert
            Assert.AreEqual(string.Format("A1:B{0}", ExcelPackage.MaxRows), validation.Address.Address);
        }
        [TestMethod]
        public void TestInsertRowsIntoVeryLongRangeWithDataValidation()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with data validation on the whole of column A except row 1
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var dvAddress = "A2:A1048576";
                var dv = wks.DataValidations.AddCustomValidation(dvAddress);

                // Check that the data validation address was set correctly
                Assert.AreEqual(dvAddress, dv.Address.Address);

                // Insert some rows into the worksheet
                wks.InsertRow(5, 3);

                // Check that the data validation rule still applies to the same range (since there's nowhere to extend it to)
                Assert.AreEqual(dvAddress, dv.Address.Address);
            }
        }
        [TestMethod]
        public void TestInsertRowsAboveVeryLongRangeWithDataValidation()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with data validation on the whole of column A except rows 1-10
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var dvAddress = "A11:A1048576";
                var dv = wks.DataValidations.AddAnyValidation(dvAddress);

                // Check that the data validation address was set correctly
                Assert.AreEqual(dvAddress, dv.Address.Address);

                // Insert 3 rows into the worksheet above the data validation
                wks.InsertRow(5, 3);

                // Check that the data validation starts lower down, but ends in the same place
                Assert.AreEqual("A14:A1048576", dv.Address.Address);
            }
        }

        [TestMethod]
        public void TestInsertRowsToPushDataValidationOffSheet()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with data validation on the last two rows of column A
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var dvAddress = "A1048575:A1048576";
                var dv = wks.DataValidations.AddCustomValidation(dvAddress);

                // Check that the data validation address was set correctly
                Assert.AreEqual(1, wks.DataValidations.Count);
                Assert.AreEqual(dvAddress, dv.Address.Address);

                // Insert enough rows into the worksheet above the data validation rule to push it off the sheet 
                wks.InsertRow(5, 10);

                // Check that the data validation rule no longer exists
                Assert.AreEqual(0, wks.DataValidations.Count);
            }
        }

        [TestMethod]
        public void TestLoadingWorksheet()
        {
            using (var p = OpenTemplatePackage("DataValidationReadTest.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual(4, ws.DataValidations.Count);
            }
        }

        [TestMethod]
        public void DataValidationAny_AllowsOperatorShouldBeFalse()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var dvAddress = "A1";
                var dv = wks.DataValidations.AddAnyValidation(dvAddress);

                Assert.IsFalse(dv.AllowsOperator);
            }
        }

        [TestMethod]
        public void DataValidationDefaults_AllowsOperatorShouldBeTrueOnCorrectTypes()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");

                var intValidation = wks.DataValidations.AddIntegerValidation("A1");
                var decimalValidation = wks.DataValidations.AddDecimalValidation("A2");
                var textLengthValidation = wks.DataValidations.AddTextLengthValidation("A3");
                var dateTimeValidation = wks.DataValidations.AddDateTimeValidation("A4");
                var timeValidation = wks.DataValidations.AddTimeValidation("A5");
                var customValidation = wks.DataValidations.AddCustomValidation("A6");

                Assert.IsTrue(intValidation.AllowsOperator);
                Assert.IsTrue(decimalValidation.AllowsOperator);
                Assert.IsTrue(textLengthValidation.AllowsOperator);
                Assert.IsTrue(dateTimeValidation.AllowsOperator);
                Assert.IsTrue(timeValidation.AllowsOperator);
                Assert.IsTrue(customValidation.AllowsOperator);
            }
        }

        [TestMethod]
        public void DataValidations_CloneShouldDeepCopy()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var validation = wks.DataValidations.AddIntegerValidation("A1");
                var clone = ((ExcelDataValidationInt)validation).GetClone();
                clone.Address = new ExcelDatavalidationAddress("A2", clone);

                Assert.AreNotEqual(validation.Address, clone.Address);
            }
        }

        [TestMethod]
        public void DataValidations_ShouldCopyAllProperties()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");

                List<ExcelDataValidation> validations = new List<ExcelDataValidation>
                {
                    (ExcelDataValidation)wks.DataValidations.AddIntegerValidation("A1"),
                    (ExcelDataValidation)wks.DataValidations.AddDecimalValidation("A2"),
                    (ExcelDataValidation)wks.DataValidations.AddTextLengthValidation("A3"),
                    (ExcelDataValidation)wks.DataValidations.AddDateTimeValidation("A4"),
                    (ExcelDataValidation)wks.DataValidations.AddTimeValidation("A5"),
                    (ExcelDataValidation)wks.DataValidations.AddCustomValidation("A6"),
                    (ExcelDataValidation)wks.DataValidations.AddAnyValidation("A7"),
                    (ExcelDataValidation)wks.DataValidations.AddListValidation("A9")
                };

                foreach (var validation in validations)
                {
                    validation.AllowBlank = true;
                    validation.Prompt = "prompt";
                    validation.PromptTitle = "promptTitle";
                    validation.Error = "error";
                    validation.ErrorTitle = "errorTitle";
                    validation.ShowInputMessage = true;
                    validation.ShowErrorMessage = true;
                    validation.ErrorStyle = ExcelDataValidationWarningStyle.information;

                    var clone = validation.GetClone();

                    Assert.AreEqual(validation.AllowBlank, clone.AllowBlank);
                    Assert.AreEqual(validation.Prompt, clone.Prompt);
                    Assert.AreEqual(validation.Error, clone.Error);
                    Assert.AreEqual(validation.ErrorTitle, clone.ErrorTitle);
                    Assert.AreEqual(validation.ShowInputMessage, clone.ShowInputMessage);
                    Assert.AreEqual(validation.ShowErrorMessage, clone.ShowErrorMessage);
                    Assert.AreEqual(validation.ErrorStyle, clone.ErrorStyle);
                }
            }
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadIMEmode()
        {
            using (var pck = OpenPackage("ImeTest.xlsx", true))
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");

                var validation = wks.DataValidations.AddCustomValidation("A1");
                validation.ShowErrorMessage = true;
                validation.Formula.ExcelFormula = "=ISTEXT(A1)";
                validation.ImeMode = ExcelDataValidationImeMode.FullKatakana;

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadIMEmodeAndWriteAgain()
        {
            using (var pck = OpenPackage("ImeTestOFF.xlsx", true))
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");

                var validation = wks.DataValidations.AddCustomValidation("A1");
                validation.ShowErrorMessage = true;
                validation.Formula.ExcelFormula = "=ISTEXT(A1)";
                validation.ImeMode = ExcelDataValidationImeMode.Off;

                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var pck2 = new ExcelPackage(stream);

                pck2.SaveAs(stream);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void DataValidations_Insert_Test()
        {
            using (var pck = OpenPackage("InsertTest.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("InsertTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A3");

                var rangeValidation2 = ws.DataValidations.AddDecimalValidation("A52");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;

                rangeValidation.Address.Address = "A1,A3";

                var list = ws.DataValidations.AddListValidation("A2");

                list.Formula.Values.Add("TestValue");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidation()
        {
            using (var pck = OpenPackage("ClearDataValidationTest.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A3");
                var rangeValidation2 = ws.DataValidations.AddIntegerValidation("A4:A6");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;


                ws.Cells["A2"].DataValidation.ClearDataValidation();

                var list = ws.DataValidations.AddListValidation("A2");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidationAndAddressChangeWithSpacedAddresses()
        {
            using (var pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A3 B5 C3 E15:E17");
                var rangeValidation2 = ws.DataValidations.AddIntegerValidation("A4:A6");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;


                ws.Cells["A2:A3"].DataValidation.ClearDataValidation();
                ws.Cells["E16 A5"].DataValidation.ClearDataValidation();

                var list = ws.DataValidations.AddListValidation("A2");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                SaveAndCleanup(pck);
            }

            using (var pck2 = OpenPackage("ClearDataValidationTestAdress.xlsx"))
            {
                var ws2 = pck2.Workbook.Worksheets[0];
                var address = ws2.DataValidations[0].Address;
                Assert.IsTrue(address.Collide(new ExcelAddressBase("A3")) == ExcelAddressBase.eAddressCollition.No);
            }
        }

        [TestMethod]
        public void ClearValidationAndAddressChangeWithSpacedAddressesViaCells()
        {
            using (var pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");

                var listValidation = ws.Cells["A1:A3 B5 C3 E15:E17"].DataValidation.AddListDataValidation();

                listValidation.Formula.Values.Add("Value1");

                var rangeValidation = ws.Cells["A4:A6"].DataValidation.AddIntegerDataValidation();

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;


                ws.Cells["A2:A3"].DataValidation.ClearDataValidation();
                ws.Cells["E16 A5"].DataValidation.ClearDataValidation();

                var list = ws.DataValidations.AddListValidation("A2");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidationOverARangeWithMultipleValidations()
        {
            using (var pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A5");
                var rangeValidation2 = ws.DataValidations.AddIntegerValidation("A6:A8");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;


                ws.Cells["A4:A7"].DataValidation.ClearDataValidation();

                var list = ws.DataValidations.AddListValidation("A4");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                Assert.AreEqual(rangeValidation.Address.Address, "A1:A3");
                Assert.AreEqual(rangeValidation2.Address.Address, "A8");

                SaveAndCleanup(pck);
            }
        }


        [TestMethod]
        public void ClearValidationOverARangeWithMultipleValidations2()
        {
            using (var pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A5");
                var rangeValidation2 = ws.DataValidations.AddIntegerValidation("A6:A8");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;

                ws.Cells["A3:A7"].DataValidation.ClearDataValidation();

                var list = ws.DataValidations.AddListValidation("A4");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                Assert.AreEqual(rangeValidation.Address.Address, "A1:A2");
                Assert.AreEqual(rangeValidation2.Address.Address, "A8");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidationOverBlockRanges()
        {
            using (var pck = OpenPackage("ClearBlockRanges.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A1:D5");
                var rangeValidation2 = ws.DataValidations.AddIntegerValidation("C6:C8");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;

                ws.Cells["A3:B7"].DataValidation.ClearDataValidation();

                var list = ws.DataValidations.AddListValidation("A4");
                var list2 = ws.DataValidations.AddListValidation("B3:B7");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");


                list2.Formula.Values.Add("Value21");
                list2.Formula.Values.Add("Value22");

                //Assert.AreEqual(rangeValidation.Address.Address, "A1:A2");
                //Assert.AreEqual(rangeValidation2.Address.Address, "A8");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void DeleteRangeOneAddressTest()
        {
            using (var pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A5");
                var rangeValidation2 = ws.DataValidations.AddIntegerValidation("A6:A8");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;

                ws.DeleteRow(2, 5);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ClearSingular()
        {
            using (var pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A9");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                ws.Cells["A9"].DataValidation.ClearDataValidation();
            }
        }

        [TestMethod]
        public void ClearSingularSpaceSeparated()
        {
            using (var pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("ClearTest");
                var rangeValidation = ws.DataValidations.AddIntegerValidation("A9 A6 B12 C50");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                ws.Cells["A9"].DataValidation.ClearDataValidation();
                ws.Cells["B12"].DataValidation.ClearDataValidation();

                Assert.AreEqual("A6 C50", rangeValidation.Address.Address);
            }
        }

        [TestMethod]
        public void RemovalOfCellsAfterBeingRemovedAndAdded()
        {
            using (var pck = OpenPackage("DataValidationsUserClearTest.xlsx", true))
            {
                var myWS = pck.Workbook.Worksheets.Add("MyWorksheet");
                var yourWS = pck.Workbook.Worksheets.Add("YourWorksheet");

                var validation = myWS.DataValidations.AddTextLengthValidation("A1:C5");

                validation.Operator = ExcelDataValidationOperator.lessThan;

                validation.Formula.Value = 10;

                myWS.Cells["B3:C6"].DataValidation.ClearDataValidation();

                var decimalVal = myWS.Cells["B3:D4"].DataValidation.AddDecimalDataValidation();
                decimalVal.Operator = ExcelDataValidationOperator.greaterThan;
                decimalVal.Formula.Value = 5;

                myWS.Cells["B1:D2"].DataValidation.ClearDataValidation();

                SaveAndCleanup(pck);
            }
        }


        [TestMethod]
        public void UserTestClear()
        {
            using (var pck = OpenPackage("DataValidationsUserClearTest.xlsx", true))
            {
                var myWS = pck.Workbook.Worksheets.Add("MyWorksheet");
                var yourWS = pck.Workbook.Worksheets.Add("YourWorksheet");

                var validation = myWS.DataValidations.AddTextLengthValidation("A1:E30");

                validation.Operator = ExcelDataValidationOperator.lessThan;

                validation.Formula.Value = 10;

                myWS.Cells["C1:D10"].DataValidation.ClearDataValidation();

                var listVal = myWS.Cells["C5:D10"].DataValidation.AddListDataValidation();

                listVal.Formula.ExcelFormula = "$C$1:$D$4";

                myWS.Cells["B1:C4"].DataValidation.ClearDataValidation();

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void FormulasWithQuotationsInExcelFormulaReadWrite()
        {
            using (var pck = OpenPackage("DV_ExcelFormulaQuotations.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("EmptyFormulaTest");
                var validation = sheet.DataValidations.AddDecimalValidation("A1");

                validation.Operator = ExcelDataValidationOperator.equal;
                validation.Formula.ExcelFormula = "\"\"\"tiger\"";

                SaveAndCleanup(pck);
                ExcelPackage readPck = OpenPackage("DV_ExcelFormulaQuotations.xlsx");
                var validationRead = readPck.Workbook.Worksheets[0].DataValidations[0];

                Assert.AreEqual("\"\"\"tiger\"", validationRead.As.DecimalValidation.Formula.ExcelFormula);
            }
        }


        [TestMethod]
        public void FormulasWithQuotationsInListReadWrite()
        {
            using (var pck = OpenPackage("DV_ListQuotations.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("EmptyFormulaTest");
                var validation = sheet.DataValidations.AddListValidation("A1");

                validation.Formula.Values.Add("\"tiger");
                validation.Formula.Values.Add("5'7\"");
                //Ensure Empty values are not read or read wrong
                validation.Formula.Values.Add("");

                SaveAndCleanup(pck);
                ExcelPackage readPck = OpenPackage("DV_ListQuotations.xlsx");
                var validationRead = readPck.Workbook.Worksheets[0].DataValidations[0];

                Assert.AreEqual("\"tiger", validationRead.As.ListValidation.Formula.Values[0]);
                Assert.AreEqual("5'7\"", validationRead.As.ListValidation.Formula.Values[1]);
                Assert.AreEqual(2, validationRead.As.ListValidation.Formula.Values.Count);
            }
        }


        [TestMethod]
        public void InsertDeleteTest()
        {
            using (var pck = OpenPackage("DV_InsertDelete.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("insertDel");
                var topValidation = sheet.DataValidations.AddDecimalValidation("H5");
                var midValidation = sheet.DataValidations.AddDecimalValidation("H10");
                var midValidation2 = sheet.DataValidations.AddDecimalValidation("G10");

                var bottomValidation = sheet.DataValidations.AddDecimalValidation("H20");

                topValidation.Operator = ExcelDataValidationOperator.equal;
                midValidation.Operator = ExcelDataValidationOperator.equal;
                midValidation2.Operator = ExcelDataValidationOperator.equal;
                bottomValidation.Operator = ExcelDataValidationOperator.equal;

                topValidation.Formula.ExcelFormula = "B1";
                midValidation.Formula.ExcelFormula = "H9";
                bottomValidation.Formula.ExcelFormula = "D11";

                sheet.InsertRow(9, 5);

                Assert.AreEqual(topValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(5, 8, 5, 8)[0]);
                Assert.AreEqual(midValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(15, 8, 15, 8)[0]);
                Assert.AreEqual(midValidation2, sheet.DataValidations._validationsRD.GetValuesFromRange(15, 7, 15, 7)[0]);
                Assert.AreEqual(bottomValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(25, 8, 25, 8)[0]);

                Assert.AreEqual("B1", topValidation.Formula.ExcelFormula);
                Assert.AreEqual("H14", midValidation.Formula.ExcelFormula);
                Assert.AreEqual("D16", bottomValidation.Formula.ExcelFormula);

                sheet.InsertRow(16, 50);
                Assert.AreEqual(midValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(15, 8, 15, 8)[0]);
                Assert.AreEqual(midValidation2, sheet.DataValidations._validationsRD.GetValuesFromRange(15, 7, 15, 7)[0]);
                Assert.AreEqual(bottomValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(75, 8, 75, 8)[0]);

                Assert.AreEqual("B1", topValidation.Formula.ExcelFormula);
                Assert.AreEqual("H14", midValidation.Formula.ExcelFormula);
                Assert.AreEqual("D66", bottomValidation.Formula.ExcelFormula);

                sheet.DeleteRow(9);

                Assert.AreEqual(topValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(5, 8, 5, 8)[0]);
                Assert.AreEqual(midValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(14, 8, 14, 8)[0]);
                Assert.AreEqual(midValidation2, sheet.DataValidations._validationsRD.GetValuesFromRange(14, 7, 14, 7)[0]);
                Assert.AreEqual(bottomValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(74, 8, 74, 8)[0]);

                Assert.AreEqual("B1", topValidation.Formula.ExcelFormula);
                Assert.AreEqual("H13", midValidation.Formula.ExcelFormula);
                Assert.AreEqual("D65", bottomValidation.Formula.ExcelFormula);

                sheet.DeleteRow(14);

                Assert.AreEqual("B1", topValidation.Formula.ExcelFormula);
                Assert.AreEqual("H13", midValidation.Formula.ExcelFormula);
                Assert.AreEqual("D64", bottomValidation.Formula.ExcelFormula);

                Assert.IsFalse(sheet.DataValidations._validationsRD.Exists(midValidation.Address._fromRow, midValidation.Address._fromCol,
                                                                           midValidation.Address._toRow, midValidation.Address._toCol));
                Assert.IsFalse(sheet.DataValidations._validationsRD.Exists(midValidation2.Address._fromRow, midValidation2.Address._fromCol,
                                                                           midValidation2.Address._toRow, midValidation2.Address._toCol));

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void InsertDeleteTestRanges()
        {
            using (var pck = OpenPackage("DV_InsertDeleteRanges.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("insertDel");
                var topValidation = sheet.DataValidations.AddListValidation("G5:G50");
                var midValidation = sheet.DataValidations.AddListValidation("H10:H25");

                topValidation.Formula.ExcelFormula = "G4:G50";
                midValidation.Formula.ExcelFormula = "H10:H25";

                sheet.InsertRow(9, 5);

                Assert.AreEqual("G4:G55", topValidation.Formula.ExcelFormula);
                Assert.AreEqual("H15:H30", midValidation.Formula.ExcelFormula);

                sheet.InsertRow(16, 50);

                Assert.AreEqual(topValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(16, 7, 20, 7)[0]);
                Assert.AreEqual(midValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(16, 8, 25, 8)[0]);

                Assert.AreEqual("G4:G105", topValidation.Formula.ExcelFormula);
                Assert.AreEqual("H15:H80", midValidation.Formula.ExcelFormula);

                sheet.DeleteRow(9);

                Assert.AreEqual(topValidation, sheet.DataValidations._validationsRD.GetValuesFromRange(104, 7, 104, 7)[0]);
                Assert.AreEqual(0, sheet.DataValidations._validationsRD.GetValuesFromRange(80, 8, 80, 8).Count);

                Assert.AreEqual("G4:G104", topValidation.Formula.ExcelFormula);
                Assert.AreEqual("H14:H79", midValidation.Formula.ExcelFormula);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void DataValidationWith2FormulasThatAreNullShouldNotThrow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("nullformulas");

                var intValidation = sheet.DataValidations.AddIntegerValidation("C2");
                intValidation.Operator = ExcelDataValidationOperator.between;

                intValidation.AllowBlank = true;

                Stream stream = new MemoryStream();
                package.SaveAs(stream);

                var readPck = new ExcelPackage(stream);

                var validation = (ExcelDataValidationInt)readPck.Workbook.Worksheets[0].DataValidations[0];
                var address = validation.Address.Address;
                Assert.AreEqual("C2", address);
                //Assert.IsNull(validation.Formula);
                //Assert.IsNull(validation.Formula2);
            }
        }

        [TestMethod]
        public void AddressContainingOwnSheetName_ShouldNotThrow()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("ExSheet");
                // add validation rules to ProblemMeta sheet
                ExcelWorksheet exSheet = package.Workbook.Worksheets.GetByName("ExSheet");
                if (exSheet != null)
                {
                    var intValidation = exSheet.DataValidations.AddIntegerValidation("ExSheet!$C$2");
                    intValidation.Operator = ExcelDataValidationOperator.equal;
                }

                Stream stream = new MemoryStream();
                package.SaveAs(stream);

                var readPck = new ExcelPackage(stream);

                var address = readPck.Workbook.Worksheets[0].DataValidations[0].Address.Address;
                Assert.AreEqual("$C$2", address);
            }
        }

        [TestMethod]
        public void MultipleAddressContainingOwnSheetName_ShouldNotThrow()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("TestSheet");
                package.Workbook.Worksheets.Add("ExtTestSheet");

                // add validation rules to ProblemMeta sheet
                ExcelWorksheet testSheet = package.Workbook.Worksheets.GetByName("TestSheet");
                ExcelWorksheet extTest = package.Workbook.Worksheets.GetByName("ExtTestSheet");

                if (testSheet != null)
                {
                    var intValidation = testSheet.DataValidations.AddIntegerValidation("TestSheet!$C$2 TestSheet!$Z$2");
                    intValidation.Operator = ExcelDataValidationOperator.equal;

                    intValidation.Formula.ExcelFormula = "ExtTestSheet!$C$5";
                }

                Stream stream = new MemoryStream();
                package.SaveAs(stream);

                var readPck = new ExcelPackage(stream);

                var validation = readPck.Workbook.Worksheets[0].DataValidations[0];
                var address = readPck.Workbook.Worksheets[0].DataValidations[0].Address.Address;
                Assert.AreEqual("$C$2,$Z$2", address);
                Assert.AreEqual("ExtTestSheet!$C$5", validation.As.IntegerValidation.Formula.ExcelFormula);
            }
        }

        [TestMethod]
        public void MultipleAddressContainingOtherSheetName_ShouldNotThrow()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("TestSheet");
                package.Workbook.Worksheets.Add("ExtTestSheet");

                // add validation rules to ProblemMeta sheet
                ExcelWorksheet testSheet = package.Workbook.Worksheets.GetByName("TestSheet");
                ExcelWorksheet extTest = package.Workbook.Worksheets.GetByName("ExtTestSheet");

                if (testSheet != null)
                {
                    //We ignore the worksheet name and only apply the addresses
                    var defaultValidation = testSheet.DataValidations.AddIntegerValidation("ExtTestSheet!$C$2 ExtTestSheet!$Z$2");
                    defaultValidation.Operator = ExcelDataValidationOperator.equal;

                    defaultValidation.Formula.ExcelFormula = "ExtTestSheet!$C$5";

                }

                Stream stream = new MemoryStream();
                package.SaveAs(stream);

                var readPck = new ExcelPackage(stream);

                var validation = readPck.Workbook.Worksheets[0].DataValidations[0];
                var address = readPck.Workbook.Worksheets[0].DataValidations[0].Address.Address;
                Assert.AreEqual("$C$2,$Z$2", address);
                Assert.AreEqual("ExtTestSheet!$C$5", validation.As.IntegerValidation.Formula.ExcelFormula);
            }
        }

        [TestMethod]
        public void OwnSheetNameApostrophe_ShouldNotThrow()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Test' She' et");
                package.Workbook.Worksheets.Add("ExtTestSheet");

                // add validation rules to ProblemMeta sheet
                ExcelWorksheet testSheet = package.Workbook.Worksheets.GetByName("Test' She' et");
                ExcelWorksheet extTest = package.Workbook.Worksheets.GetByName("ExtTestSheet");

                if (testSheet != null)
                {
                    //We ignore the worksheet name and only apply the addresses
                    var defaultValidation = testSheet.DataValidations.AddIntegerValidation("Test' She' et!$C$2 Test' She' et!$Z$5");
                    defaultValidation.Operator = ExcelDataValidationOperator.equal;

                    defaultValidation.Formula.ExcelFormula = "ExtTestSheet!$C$5";
                }

                Stream stream = new MemoryStream();
                package.SaveAs(stream);

                var readPck = new ExcelPackage(stream);

                var validation = readPck.Workbook.Worksheets[0].DataValidations[0];
                var address = readPck.Workbook.Worksheets[0].DataValidations[0].Address.Address;
                Assert.AreEqual("$C$2,$Z$5", address);
                Assert.AreEqual("ExtTestSheet!$C$5", validation.As.IntegerValidation.Formula.ExcelFormula);
            }
        }

        [TestMethod]
        public void HandleReadingOfOldDataValidationFiles()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("oldWorksheetMock");
                package.Workbook.Worksheets.Add("extSheet");

                string oldValidations = "";
                string attributes = "<dataValidation allowBlank=\"true\" errorStyle=\"stop\" operator=\"equal\" showDropDown=\"false\" showErrorMessage=\"true\" showInputMessage=\"false\"";
                for (int i = 0; i < 5; i++) 
                {
                    char nextChar = (char)('j' + i);
                    oldValidations += attributes + $" sqref=\"{nextChar}{i+1}\"" + " type=\"list\">";
                    oldValidations += "<formula1>'extSheet'!$A:$A</formula1>";
                    oldValidations += "<formula2>0</formula2>";
                    oldValidations += "</dataValidation>";
                }

                string xmlMock = "<root><dataValidations count=\"5\">" + oldValidations + "</dataValidations><printOptions headings=\"false\" gridLines=\"false\" gridLinesSet=\"true\" horizontalCentered=\"false\" verticalCentered=\"false\"/>";
                xmlMock += "</root>";

                var reader = new StringReader(xmlMock);

                var xr = XmlReader.Create(reader, new XmlReaderSettings()
                {
                    ConformanceLevel = ConformanceLevel.Fragment,
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreWhitespace = true,
                });

                xr.Read();
                xr.Read();

                sheet.DataValidations.ReadDataValidations(xr);

                Assert.AreEqual(5, sheet.DataValidations.Count);
                Assert.AreEqual("J1", sheet.DataValidations[0].Address.Address);
                Assert.AreEqual("K2", sheet.DataValidations[1].Address.Address);
                Assert.AreEqual("L3", sheet.DataValidations[2].Address.Address);
                Assert.AreEqual("M4", sheet.DataValidations[3].Address.Address);
                Assert.AreEqual("N5", sheet.DataValidations[4].Address.Address);
                Assert.AreEqual("printOptions", xr.LocalName);
            }
        }

        [TestMethod]
        public void EnsureEmptyFormulasCanBeRead()
        {
            using(var package = OpenPackage("EmptyDataValidationsTest.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("emptyValidations");

                var validation = sheet.Cells["A1:B5"].DataValidation.AddDecimalDataValidation();

                validation.Operator = ExcelDataValidationOperator.between;
                validation.AllowBlank = true;

                SaveAndCleanup(package);

                var readPackage = OpenPackage("EmptyDataValidationsTest.xlsx");

                var readSheet = readPackage.Workbook.Worksheets.FirstOrDefault();

                var readValidation = readSheet.DataValidations[0].As.DecimalValidation;

                Assert.IsNull(readValidation.Formula.Value);
                Assert.IsNull(readValidation.Formula2.Value);
            }
        }

        [TestMethod]
        public void Issue1714()
        {
            using var p = OpenTemplatePackage("DataValidationIssue1714.xlsx");
            foreach (var v in p.Workbook.Worksheets[0].DataValidations)
            {
                Assert.IsTrue(v.AllowBlank);
                Assert.IsFalse(v.ShowInputMessage);
                Assert.IsTrue(v.ShowErrorMessage);
            }
        }
    }
}