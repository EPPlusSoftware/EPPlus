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
using OfficeOpenXml.DataValidation.Contracts;
using System;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ListDataValidationTests : ValidationTestBase
    {
        private IExcelDataValidationList _validation;

        [TestInitialize]
        public void Setup()
        {
            SetupTestData();
            _validation = _sheet.Workbook.Worksheets[1].DataValidations.AddListValidation("A1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            CleanupTestData();
        }

        [TestMethod]
        public void ListDataValidation_FormulaIsSet()
        {
            Assert.IsNotNull(_validation.Formula);
        }

        [TestMethod]
        public void ListDataValidation_CanAssignFormula()
        {
            _validation.Formula.ExcelFormula = "abc!A2";
            Assert.AreEqual("abc!A2", _validation.Formula.ExcelFormula);
        }
        [TestMethod]
        public void ListDataValidation_CanAssignDefinedName()
        {
            _validation.Formula.ExcelFormula = "ListData";
            Assert.AreEqual("ListData", _validation.Formula.ExcelFormula);
        }

        [TestMethod]
        public void ListDataValidation_WhenOneItemIsAddedCountIs1()
        {
            // Act
            _validation.Formula.Values.Add("test");

            // Assert
            Assert.AreEqual(1, _validation.Formula.Values.Count);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ListDataValidation_ShouldThrowWhenNoFormulaOrValueIsSet()
        {
            _validation.Validate();
        }

        [TestMethod]
        public void ListDataValidation_ShowErrorMessageIsSet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("list formula");

                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet2.Cells["A1"].Value = "A";
                sheet2.Cells["A2"].Value = "B";
                sheet2.Cells["A3"].Value = "C";

                // add a validation and set values
                var validation = sheet.DataValidations.AddListValidation("A1");
                // Alternatively:
                // var validation = sheet.Cells["A1"].DataValidation.AddListDataValidation();
                validation.ShowErrorMessage = true;
                validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                validation.ErrorTitle = "An invalid value was entered";
                validation.Error = "Select a value from the list";
                validation.Formula.ExcelFormula = "Sheet2!A1:A3";

                Assert.IsTrue(validation.ShowErrorMessage.Value);
            }
        }

        [TestMethod]
        public void ListDataValidationExt_ShowDropDownIsSet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("list formula");

                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet2.Cells["A1"].Value = "A";
                sheet2.Cells["A2"].Value = "B";
                sheet2.Cells["A3"].Value = "C";

                // add a validation and set values
                var validation = sheet.DataValidations.AddListValidation("A1");
                // Alternatively:
                // var validation = sheet.Cells["A1"].DataValidation.AddListDataValidation();
                validation.HideDropDown = true;
                validation.ShowErrorMessage = true;
                validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                validation.ErrorTitle = "An invalid value was entered";
                validation.Error = "Select a value from the list";
                validation.Formula.ExcelFormula = "Sheet2!A1:A3";

                // refresh the data validation
                validation = sheet.DataValidations.Find(x => x.Uid == validation.Uid).As.ListValidation;

                Assert.IsTrue(validation.HideDropDown.Value);
                var v = validation as ExcelDataValidationList;
                var attributeValue = v.HideDropDown.Value;
                Assert.IsTrue(attributeValue);
            }
        }

        [TestMethod]
        public void ListDataValidation_ShowDropDownIsSet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("list formula");
                sheet.Cells["A1"].Value = "A";
                sheet.Cells["A2"].Value = "B";
                sheet.Cells["A3"].Value = "C";

                // add a validation and set values
                var validation = sheet.DataValidations.AddListValidation("B1");
                validation.HideDropDown = true;
                validation.ShowErrorMessage = true;
                validation.Formula.ExcelFormula = "A1:A3";

                Assert.IsTrue(validation.HideDropDown.Value);
                var v = validation as ExcelDataValidationList;
                var attributeValue = v.HideDropDown.Value;
                Assert.IsTrue(attributeValue);
            }
        }

        [TestMethod]
        public void ListDataValidation_AllowsOperatorShouldBeFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("list operator");

                // add a validation and set values
                var validation = sheet.DataValidations.AddListValidation("A1");

                Assert.IsFalse(validation.AllowsOperator);
            }
        }
    }
}
