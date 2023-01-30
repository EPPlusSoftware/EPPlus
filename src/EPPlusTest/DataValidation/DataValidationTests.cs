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
using OfficeOpenXml.Utils;
using System;
using System.IO;
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

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsEqualAndFormula1IsEmpty()
        {
            var validations = _sheet.DataValidations.AddIntegerValidation("A1");
            validations.Operator = ExcelDataValidationOperator.equal;
            validations.Validate();
        }

        [TestMethod]
        public void DataValidations_ShouldReadWriteTypes()
        {
            var P = new ExcelPackage(new MemoryStream());
            var sheet = P.Workbook.Worksheets.Add("NewSheet");

            sheet.DataValidations.AddAnyValidation("A1");
            sheet.DataValidations.AddIntegerValidation("A2");
            sheet.DataValidations.AddDecimalValidation("A3");
            sheet.DataValidations.AddListValidation("A4");
            sheet.DataValidations.AddTextLengthValidation("A5");
            sheet.DataValidations.AddDateTimeValidation("A6");
            sheet.DataValidations.AddTimeValidation("A7");
            sheet.DataValidations.AddCustomValidation("A8");

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
        public void DataValidations_ShouldSetOperatorFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "greaterThanOrEqual", "1");
            // Act
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), "A1", _sheet.Name);
            // Assert
            Assert.AreEqual(ExcelDataValidationOperator.greaterThanOrEqual, validation.Operator);
        }

        [TestMethod]
        public void DataValidations_ShouldSetShowErrorMessageFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", true, false);
            // Act
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), "A1", _sheet.Name);
            // Assert
            Assert.IsTrue(validation.ShowErrorMessage ?? false);
        }

        [TestMethod]
        public void DataValidations_ShouldSetShowInputMessageFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", false, true);
            // Act
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), "A1", _sheet.Name);
            // Assert
            Assert.IsTrue(validation.ShowInputMessage ?? false);
        }

        [TestMethod]
        public void DataValidations_ShouldSetPromptFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), "A1", _sheet.Name);
            // Assert
            Assert.AreEqual("Prompt", validation.Prompt);
        }

        [TestMethod]
        public void DataValidations_ShouldSetPromptTitleFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), "A1", _sheet.Name);
            // Assert
            Assert.AreEqual("PromptTitle", validation.PromptTitle);
        }

        [TestMethod]
        public void DataValidations_ShouldSetErrorFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), "A1", _sheet.Name);
            // Assert
            Assert.AreEqual("Error", validation.Error);
        }

        [TestMethod]
        public void DataValidations_ShouldReadErrorFromExistingXml()
        {
            // Arrange
            XmlReader xr = new XmlNodeReader(LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle"));
            xr.ReadUntil("dataValidation");
            // Act
            var validation = new ExcelDataValidationInt(xr);
            // Assert
            Assert.AreEqual("Error", validation.Error);
        }


        [TestMethod]
        public void DataValidations_ShouldSetErrorTitleFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), "A1", _sheet.Name);
            // Assert
            Assert.AreEqual("ErrorTitle", validation.ErrorTitle);
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

        [TestMethod]
        public void DataValidations_ShouldNotThrowIfAllowBlankIsSet()
        {
            var validation = _sheet.DataValidations.AddListValidation("A1");
            validation.AllowBlank = true;
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
    }
}
