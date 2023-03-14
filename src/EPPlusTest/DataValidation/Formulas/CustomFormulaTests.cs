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
    public class CustomFormulaTests : ValidationTestBase
    {
        [TestMethod]
        public void FormulaIsReadInConstructor()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("CustomTest");

            var validationOrig = sheet.DataValidations.AddCustomValidation("A1");
            validationOrig.Formula.ExcelFormula = "A1";

            var validation = ReadTValidation<ExcelDataValidationCustom>(package);

            Assert.AreEqual("A1", validation.Formula.ExcelFormula);
        }

        [TestMethod]
        public void FormulaIsValidlyParsed()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("IntegerTest");

            var validationAmpersand = sheet.DataValidations.AddCustomValidation("A1");

            sheet.Cells["B1"].Value = "EP";
            sheet.Cells["C1"].Value = "Plus";

            validationAmpersand.Formula.ExcelFormula = "=B1&C1=A1";
            validationAmpersand.ShowErrorMessage = true;

            var lessThan = sheet.DataValidations.AddCustomValidation("A2");

            sheet.Cells["B2"].Value = 1;
            sheet.Cells["C2"].Value = 2;

            lessThan.Formula.ExcelFormula = "B2<C2";
            lessThan.ShowErrorMessage = true;

            var largerThan = sheet.DataValidations.AddCustomValidation("A3");
            largerThan.Formula.ExcelFormula = "C2>B2";
            largerThan.ShowErrorMessage = true;

            MemoryStream stream = new MemoryStream();
            package.SaveAs("C:\\epplusTest\\Workbooks/escapeTest.xlsx");
        }
    }
}
