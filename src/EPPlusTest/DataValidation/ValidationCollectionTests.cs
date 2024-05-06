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
using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ValidationCollectionTests : ValidationTestBase
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

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenAddressIsNullOrEmpty()
        {
            // Act
            _sheet.DataValidations.AddDecimalValidation(string.Empty);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidations.AddDecimalValidation("A1");
            _sheet.DataValidations.AddDecimalValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddInteger_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidations.AddIntegerValidation("A1");
            _sheet.DataValidations.AddIntegerValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddTextLength_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidations.AddTextLengthValidation("A1");
            _sheet.DataValidations.AddTextLengthValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddDateTime_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A1");
        }

        [TestMethod]
        public void ExcelDataValidationCollection_Index_ShouldReturnItemAtIndex()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");
            _sheet.DataValidations.AddDateTimeValidation("B1");

            // Act
            var result = _sheet.DataValidations[1];

            // Assert
            Assert.AreEqual("A2", result.Address.Address);
        }

        [TestMethod]
        public void ExcelDataValidationCollection_FindAll_ShouldReturnValidationInColumnAonly()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");
            _sheet.DataValidations.AddDateTimeValidation("B1");

            // Act
            var result = _sheet.DataValidations.FindAll(x => x.Address.Address.StartsWith("A"));

            // Assert
            Assert.AreEqual(2, result.Count());

        }

        [TestMethod]
        public void ExcelDataValidationCollection_Find_ShouldReturnFirstMatchOnly()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");

            // Act
            var result = _sheet.DataValidations.Find(x => x.Address.Address.StartsWith("A"));

            // Assert
            Assert.AreEqual("A1", result.Address.Address);

        }

        [TestMethod]
        public void ExcelDataValidationCollection_Clear_ShouldBeEmpty()
        {
            // Arrange
            var v = _sheet.DataValidations.AddDateTimeValidation("A1");

            // Act
            _sheet.DataValidations.Clear();

            // Assert
            Assert.AreEqual(0, _sheet.DataValidations.Count);

        }

        [TestMethod]
        public void ExcelDataValidationCollection_Clear_ShouldEmptyValidationsRDToo()
        {
            // Arrange
            var v = _sheet.DataValidations.AddDateTimeValidation("A1");

            // Act
            _sheet.DataValidations.Clear();

            // Assert
            Assert.AreEqual(0, _sheet.DataValidations.Count);

            var validation = _sheet.DataValidations.AddCustomValidation("A1");
            _sheet.DataValidations.Clear();
        }

        [TestMethod]
        public void ExcelDataValidationCollection_RemoveShouldRemoveFromRD()
        {
            _sheet.DataValidations.Clear();

            // Arrange
            var v = _sheet.DataValidations.AddDateTimeValidation("A1:A3");
            var v2 = _sheet.DataValidations.AddDecimalValidation("B1:B3");
            var v3 = _sheet.DataValidations.AddAnyValidation("A4:C4");

            _sheet.DataValidations.Remove(v2);

            var addresses = _sheet.DataValidations._validationsRD._addresses;

            Assert.AreEqual(1, addresses[2].Count);

            var validation = _sheet.DataValidations.AddCustomValidation("B2");
            _sheet.DataValidations.Clear();
        }


        [TestMethod]
        public void ExcelDataValidationCollection_ExtLst_Clear_ShouldBeEmpty()
        {
            // Arrange
            var sheet2 = _package.Workbook.Worksheets.Add("Sheet2");
            var v = _sheet.DataValidations.AddListValidation("A1");
            v.Formula.ExcelFormula = "Sheet2!A1:A2";

            // Act
            _sheet.DataValidations.Clear();

            // Assert
            Assert.AreEqual(0, _sheet.DataValidations.Count);

        }

        [TestMethod]
        public void ExcelDataValidationCollection_RemoveAll_ShouldRemoveMatchingEntries()
        {
            // Arrange
            _sheet.DataValidations.AddIntegerValidation("A1");
            _sheet.DataValidations.AddIntegerValidation("A2");
            _sheet.DataValidations.AddIntegerValidation("B1");

            // Act
            _sheet.DataValidations.RemoveAll(x => x.Address.Address.StartsWith("B"));

            // Assert
            Assert.AreEqual(2, _sheet.DataValidations.Count);
        }
    }
}
