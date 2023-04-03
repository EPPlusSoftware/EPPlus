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
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

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

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsEqualAndFormula1IsEmpty()
        {
            var validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.Operator = ExcelDataValidationOperator.equal;
            validation.Validate();
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadAllValidOperatorsOnAllTypes()
        {

        }

        public void TestTypeOperator(ExcelDataValidation type)
        {

        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException)), Ignore]
        public void TestRangeAddMultipleTryAddingAfterShouldThrow()
        {
            ExcelPackage pck = new ExcelPackage("C:\\epplusTest\\Workbooks\\ValidationRangeTestMany.xlsx");

            var validations = pck.Workbook.Worksheets[0].DataValidations;

            validations.AddIntegerValidation("C8");
        }

        [TestMethod, Ignore]
        public void TestRangeAddMultipleTryAddingAfterShouldNotThrow()
        {
            ExcelPackage pck = new ExcelPackage("C:\\epplusTest\\Workbooks\\ValidationRangeTestMany.xlsx");

            var validations = pck.Workbook.Worksheets[0].DataValidations;

            validations.AddIntegerValidation("Z8");
        }


        [TestMethod, Ignore]
        public void TestRangeAddsMultipleInbetweenInstances()
        {
            ExcelPackage pck = new ExcelPackage("C:\\epplusTest\\Workbooks\\ValidationRangeTestMany.xlsx");

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

        [TestMethod, Ignore]
        public void TestRangeAddsSingularInstance()
        {
            ExcelPackage pck = new ExcelPackage("C:\\epplusTest\\Workbooks\\ValidationRangeTest.xlsx");

            //pck.Workbook.Worksheets.Add("RangeTest");

            var validations = pck.Workbook.Worksheets[0].DataValidations;

            StringBuilder sb = new StringBuilder();

            //Ensure all addresses exist in _validationsRD
            for(int i = 0; i< validations.Count; i++) 
            {
                if(validations[i].Address.Addresses != null)
                {
                    var addresses = validations[i].Address.Addresses;

                    for (int j = 0; j < validations[i].Address.Addresses.Count; j++)
                    {
                        if(!validations._validationsRD.Exists(addresses[j]._fromRow, addresses[j]._fromCol, addresses[j]._toRow, addresses[j]._toCol))
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

            Assert.AreEqual("",sb.ToString());
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
            var validation = _sheet.DataValidations.AddListValidation("A1");
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
            using (var p = OpenTemplatePackage("DataValidationTest.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual(3, ws.DataValidations.Count);
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
                clone.Address = new ExcelAddress("A2");

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
    }
}
