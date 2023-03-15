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
using System.IO;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class IntegerFormulaTests : ValidationTestBase
    {
        [TestMethod]
        public void ValueIsRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("IntegerTest");

            var validationOrig = sheet.DataValidations.AddIntegerValidation("A1");

            validationOrig.Formula.Value = 12;
            validationOrig.Operator = ExcelDataValidationOperator.lessThanOrEqual;

            var validation = ReadTValidation<ExcelDataValidationInt>(package);

            Assert.AreEqual(12, validation.Formula.Value);
        }

        [TestMethod]
        public void ExcelFormulaIsRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("IntegerTest");

            var validationOrig = sheet.DataValidations.AddIntegerValidation("A1");

            validationOrig.Formula.ExcelFormula = "D1";
            validationOrig.Operator = ExcelDataValidationOperator.lessThanOrEqual;

            var validation = ReadTValidation<ExcelDataValidationInt>(package);

            Assert.AreEqual("D1", validation.Formula.ExcelFormula);
        }

        [TestMethod]
        public void FormulaSpecialSignsAreWrittenAndRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("IntegerTest");

            var lessThan = sheet.DataValidations.AddIntegerValidation("A1");
            lessThan.Operator = ExcelDataValidationOperator.equal;

            sheet.Cells["B1"].Value = 1;

            lessThan.Formula.ExcelFormula = "=B1<5";
            lessThan.ShowErrorMessage= true;


            var greaterThan = sheet.DataValidations.AddIntegerValidation("A2");

            sheet.Cells["B2"].Value = 6;

            greaterThan.Formula.ExcelFormula = "=B1>5";
            greaterThan.ShowErrorMessage = true;

            greaterThan.Operator = ExcelDataValidationOperator.equal;

            MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);

            var loadedpkg = new ExcelPackage(stream);
            var loadedSheet = loadedpkg.Workbook.Worksheets[0];

            var validations = loadedSheet.DataValidations;

            Assert.AreEqual(((ExcelDataValidationInt)validations[0]).Formula.ExcelFormula, "=B1<5");
            Assert.AreEqual(((ExcelDataValidationInt)validations[1]).Formula.ExcelFormula, "=B1>5");
        }
    }
}
