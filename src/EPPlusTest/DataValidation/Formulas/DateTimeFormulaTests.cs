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
using System;
using System.IO;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class DateTimeFormulaTests : ValidationTestBase
    {
        [TestMethod]
        public void FormulaValueIsRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("DateTest");

            var validationOrig = sheet.DataValidations.AddDateTimeValidation("A1");

            var date = DateTime.Parse("2011-01-08");

            validationOrig.Formula.Value = date;
            validationOrig.Operator = ExcelDataValidationOperator.lessThanOrEqual;

            var validation = ReadTValidation<ExcelDataValidationDateTime>(package);

            Assert.AreEqual(date, validation.Formula.Value);
        }

        [TestMethod]
        public void ExcelFormulaValueIsRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("DateTest");

            var validationOrig = sheet.DataValidations.AddDateTimeValidation("A1");
            validationOrig.Formula.ExcelFormula = "A1";
            validationOrig.Operator = ExcelDataValidationOperator.lessThanOrEqual;

            var validation = ReadTValidation<ExcelDataValidationDateTime>(package);
            Assert.AreEqual("A1", validation.Formula.ExcelFormula);
        }

        [TestMethod]
        public void ExcelFormulaSetToValueInsteadOfAddressIsRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("DateTest");

            var validationOrig = sheet.DataValidations.AddDateTimeValidation("A1");

            var date = DateTime.Parse("2011-01-08");
            var dateString = date.ToOADate().ToString();

            validationOrig.Formula.ExcelFormula = dateString;
            validationOrig.Operator = ExcelDataValidationOperator.lessThanOrEqual;

            var validation = ReadTValidation<ExcelDataValidationDateTime>(package);

            Assert.AreEqual(date, validation.Formula.Value);
        }

        [TestMethod]
        public void FormulaSpecialSignsAreWrittenAndRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("DateTest");

            var lessThan = sheet.DataValidations.AddDateTimeValidation("A1");
            lessThan.Operator = ExcelDataValidationOperator.equal;

            DateTime dateTime = DateTime.Parse("5/1/2008 8:30:52 AM", System.Globalization.CultureInfo.InvariantCulture);

            lessThan.Formula.Value = dateTime;

            sheet.Cells["B1"].Value = DateTime.Parse("3/15/2023 12:01:51", System.Globalization.CultureInfo.InvariantCulture);

            lessThan.Formula.ExcelFormula = $"B1<{lessThan.Formula.Value}";
            lessThan.ShowErrorMessage = true;

            var greaterThan = sheet.DataValidations.AddDateTimeValidation("A2");

            greaterThan.Formula.ExcelFormula = $"=B1>\"{dateTime}\"";
            greaterThan.ShowErrorMessage = true;

            greaterThan.Operator = ExcelDataValidationOperator.equal;

            MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);

            var loadedpkg = new ExcelPackage(stream);
            var loadedSheet = loadedpkg.Workbook.Worksheets[0];

            var validations = loadedSheet.DataValidations;

            Assert.AreEqual(((ExcelDataValidationDateTime)validations[0]).Formula.ExcelFormula, $"B1<{dateTime}");
            Assert.AreEqual(((ExcelDataValidationDateTime)validations[1]).Formula.ExcelFormula, $"=B1>\"{dateTime}\"");
        }
    }
}
