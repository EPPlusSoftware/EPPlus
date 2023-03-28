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
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class ListFormulaTests : ValidationTestBase
    {
        [TestMethod]
        public void ValuesAreReadExcelFormula()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("ListTest");

            var validationOrig = sheet.DataValidations.AddListValidation("A1");

            validationOrig.Formula.ExcelFormula = "\"3,42\"";

            var validation = ReadTValidation<ExcelDataValidationList>(package);

            Assert.AreEqual("3", validation.Formula.Values[0]);
            Assert.AreEqual("42", validation.Formula.Values[1]);
        }

        [TestMethod]
        public void ValuesAreRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("ListTest");

            var validationOrig = sheet.DataValidations.AddListValidation("A1");

            validationOrig.Formula.Values.Add("5");
            validationOrig.Formula.Values.Add("15");


            var validation = ReadTValidation<ExcelDataValidationList>(package);

            CollectionAssert.AreEquivalent(new List<string> { "5", "15" }, (ICollection)validation.Formula.Values);
        }

        [TestMethod]
        public void ExcelFormulaIsRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("ListTest");

            var validationOrig = sheet.DataValidations.AddListValidation("A1");

            validationOrig.Formula.ExcelFormula = "D1";

            var validation = ReadTValidation<ExcelDataValidationList>(package);

            Assert.AreEqual("D1", validation.Formula.ExcelFormula);
        }

        [TestMethod]
        public void FormulaSpecialSignsAreWrittenAndRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("DecimalTest");

            var lessThan = sheet.DataValidations.AddListValidation("A1");

            sheet.Cells["B1"].Value = "EP";
            sheet.Cells["B2"].Value = "Plus";

            lessThan.Formula.ExcelFormula = "\"1<5,B1&B2,2>4\"";
            lessThan.HideDropDown = false;

            lessThan.ShowErrorMessage = true;


            var greaterThan = sheet.DataValidations.AddListValidation("A2");

            sheet.Cells["B2"].Value = 6;
            greaterThan.Formula.ExcelFormula = "\"HulaBal00,-?53&<>/\\\'#¤$%||,123456789\"";

            greaterThan.ShowErrorMessage = true;
            greaterThan.HideDropDown = false;

            MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);

            var loadedpkg = new ExcelPackage(stream);
            var loadedSheet = loadedpkg.Workbook.Worksheets[0];

            var validations = loadedSheet.DataValidations;

            Assert.AreEqual(((ExcelDataValidationList)validations[0]).Formula.Values[0], "1<5");
            Assert.AreEqual(((ExcelDataValidationList)validations[0]).Formula.Values[1], "B1&B2");
            Assert.AreEqual(((ExcelDataValidationList)validations[0]).Formula.Values[2], "2>4");

            Assert.AreEqual(((ExcelDataValidationList)validations[1]).Formula.Values[0], "HulaBal00");
            Assert.AreEqual(((ExcelDataValidationList)validations[1]).Formula.Values[1], "-?53&<>/\\'#¤$%||");
            Assert.AreEqual(((ExcelDataValidationList)validations[1]).Formula.Values[2], "123456789");
        }
    }
}
