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
using System.Linq;

namespace EPPlusTest.DataValidation.IntegrationTests
{
    /// <summary>
    /// Remove the Ignore attributes from the testmethods if you want to run any of these tests
    /// </summary>
    [TestClass]
    public class IntegrationTests : TestBase
    {
        static ExcelPackage _package;
        private ExcelPackage _unitTestPackage;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _package = OpenPackage("DatavalidationIntegrationTests.xlsx", true);
        }


        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_package);
        }

        [TestInitialize]
        public void TestInitialize()
        {
            _unitTestPackage = new ExcelPackage();
        }

        [TestCleanup]
        public void TestCleanup()
        {
            _unitTestPackage.Dispose();
        }

        [TestMethod]
        public void DataValidations_AddOneValidationOfTypeWhole()
        {
            var ws = _package.Workbook.Worksheets.Add("AddOneValidationOfTypeWhole");
            ws.Cells["B1"].Value = 2;
            var validation = ws.DataValidations.AddIntegerValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.ErrorTitle = "Invalid value was entered";
            validation.PromptTitle = "Enter value here";
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            //validation.Value.Value = 3;
            validation.Formula.ExcelFormula = "B1";
        }
        [TestMethod]
        public void DataValidations_AddOneValidationOfTypeDecimal()
        {
            var ws = _package.Workbook.Worksheets.Add("AddOneValidationOfTypeDecimal");
            var validation = ws.DataValidations.AddDecimalValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.ErrorTitle = "Invalid value was entered";
            validation.Error = "Value must be greater than 1.4";
            validation.PromptTitle = "Enter value here";
            validation.Prompt = "Enter a value that is greater than 1.4";
            validation.ShowInputMessage = true;
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            validation.Formula.Value = 1.4;
        }

        [TestMethod]
        public void DataValidations_AddOneValidationOfTypeListOfTypeList()
        {
            var ws = _package.Workbook.Worksheets.Add("AddOneValidationOfTypeList");
            var validation = ws.DataValidations.AddListValidation("A:A");
            validation.ShowErrorMessage = true;
            validation.ShowInputMessage = true;
            validation.Formula.Values.Add("1");
            validation.Formula.Values.Add("2");
            validation.Formula.Values.Add("3");
            validation.Validate();
        }

        [TestMethod]
        public void DataValidations_AddOneValidationOfTypeListOfTypeTime()
        {
            var ws = _package.Workbook.Worksheets.Add("AddOneValidationOfTypeTime");
            var validation = ws.DataValidations.AddTimeValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ShowInputMessage = true;
            validation.Formula.Value.Hour = 14;
            validation.Formula.Value.Minute = 30;
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            validation.Prompt = "Enter a time greater than 14:30";
            validation.Error = "Invalid time was entered";
            validation.Validate();
        }

        [TestMethod]
        public void ShouldMoveIntValidationToExtListWhenReferringOtherWorksheet()
        {
            var sheet1 = _unitTestPackage.Workbook.Worksheets.Add("extlist_sheet1");
            var sheet2 = _unitTestPackage.Workbook.Worksheets.Add("extlist_sheet2");

            var v = sheet1.Cells["A1"].DataValidation.AddIntegerDataValidation();
            v.Formula.ExcelFormula = "extlist_sheet2!A1";
            v.Formula2.ExcelFormula = "B2";
            v.Operator = ExcelDataValidationOperator.between;
            v.ShowErrorMessage = true;
            v.ShowInputMessage = true;
            v.AllowBlank = true;

            sheet2.Cells["A1"].Value = 1;
            sheet1.Cells["A2"].Value = 3;

            SaveWorkbook("MoveToExtLstInt.xlsx", _unitTestPackage);
        }

        [TestMethod]
        public void ShouldMoveListValidationToExtListWhenReferringOtherWorksheet()
        {
            var sheet1 = _unitTestPackage.Workbook.Worksheets.Add("extlist_sheet1");
            var sheet2 = _unitTestPackage.Workbook.Worksheets.Add("extlist_sheet2");

            var v = sheet1.Cells["A1"].DataValidation.AddListDataValidation();
            v.ShowErrorMessage = true;
            v.ShowInputMessage = true;
            v.AllowBlank = true;
            v.Formula.ExcelFormula = "extlist_sheet2!A1:A2";

            sheet2.Cells["A1"].Value = "option1";
            sheet2.Cells["A2"].Value = "option2";

            SaveWorkbook("MoveToExtLst.xlsx", _unitTestPackage);
        }

        [TestMethod]
        public void ShouldMoveCustomValidationToExtListWhenReferringOtherWorksheet()
        {
            var sheet1 = _unitTestPackage.Workbook.Worksheets.Add("Sheet1");
            var sheet2 = _unitTestPackage.Workbook.Worksheets.Add("Sheet2");

            sheet1.Cells["A1"].Value = "Bar";
            sheet2.Cells["A1"].Value = "Foo";
            var v = sheet1.Cells["A1"].DataValidation.AddCustomDataValidation();
            v.ShowErrorMessage = true;
            v.ShowInputMessage = true;
            v.AllowBlank = false;
            v.Formula.ExcelFormula = "IF(AND(Sheet2!A1=\"Foo\",A1=\"Bar\"),TRUE,FALSE)";


            SaveWorkbook("MoveToExtLst_Custom.xlsx", _unitTestPackage);
        }

        [TestMethod]
        public void ShoulAddListValidationOnSameWorksheet()
        {
            var sheet1 = _package.Workbook.Worksheets.Add("extlist_sheet1");

            var v = sheet1.Cells["A1"].DataValidation.AddListDataValidation();
            v.Formula.ExcelFormula = "B1:B2";
            v.ShowErrorMessage = true;
            v.ShowInputMessage = true;
            v.AllowBlank = true;

            sheet1.Cells["B1"].Value = "option1";
            sheet1.Cells["B2"].Value = "option2";

            SaveWorkbook("DvSameWorksheet.xlsx", _unitTestPackage);
        }

        [TestMethod]
        public void ShouldMoveListValidationToExtListAndBack()
        {
            var sheet1 = _unitTestPackage.Workbook.Worksheets.Add("extlist_sheet1");
            var sheet2 = _unitTestPackage.Workbook.Worksheets.Add("extlist_sheet2");

            var v = sheet1.Cells["A1"].DataValidation.AddListDataValidation();
            v.ShowErrorMessage = true;
            v.ShowInputMessage = true;
            v.AllowBlank = true;
            v.Formula.ExcelFormula = "extlist_sheet2!A1:A2";
            v.Formula.ExcelFormula = "B1:B2";

            sheet1.Cells["B1"].Value = "option1";
            sheet1.Cells["B2"].Value = "option2";

            SaveWorkbook("MoveToExtLst.xlsx", _unitTestPackage);
        }

        [TestMethod]
        public void RemoveDataValidation()
        {
            using (var p1 = new ExcelPackage())
            {
                var sheet = p1.Workbook.Worksheets.Add("test");
                var validation = sheet.DataValidations.AddIntegerValidation("A1");
                validation.Formula.Value = 1;
                validation.Formula2.Value = 2;
                validation.ShowErrorMessage = true;
                validation.Error = "Error!";

                p1.Save();
                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    sheet = p2.Workbook.Worksheets.First();
                    var dv = sheet.DataValidations.First();
                    sheet.DataValidations.Remove(dv);
                    SaveWorkbook("RemoveDataValidation.xlsx", p2);
                }
            }
        }
    }
}
