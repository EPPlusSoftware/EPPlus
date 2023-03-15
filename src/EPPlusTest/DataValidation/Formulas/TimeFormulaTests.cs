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
using OfficeOpenXml.Style;
using System.IO;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class TimeFormulaTests : ValidationTestBase
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
        public void ValueIsRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("TimeTest");

            var validationOrig = sheet.DataValidations.AddTimeValidation("A1");

            validationOrig.Formula.Value.Hour = 14;
            validationOrig.Formula.Value.Minute = 30;
            validationOrig.Formula.Value.Second = 42;

            validationOrig.Operator = ExcelDataValidationOperator.lessThanOrEqual;

            var validation = ReadTValidation<ExcelDataValidationTime>(package);

            Assert.AreEqual(validationOrig.Formula.Value.Hour, validation.Formula.Value.Hour);
            Assert.AreEqual(validationOrig.Formula.Value.Minute, validation.Formula.Value.Minute);
            Assert.AreEqual(validationOrig.Formula.Value.Second, validation.Formula.Value.Second);
        }

        [TestMethod]
        public void ExcelFormulaIsRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("TimeTest");

            var validationOrig = sheet.DataValidations.AddTimeValidation("A1");

            validationOrig.Formula.ExcelFormula = "D1";

            validationOrig.Operator = ExcelDataValidationOperator.lessThanOrEqual;

            var validation = ReadTValidation<ExcelDataValidationTime>(package);

            Assert.AreEqual("D1", validation.Formula.ExcelFormula);
        }

        [TestMethod]
        public void FormulaSpecialSignsAreWrittenAndRead()
        {
            var package = new ExcelPackage(new MemoryStream());
            var sheet = package.Workbook.Worksheets.Add("TimeTest");

            var lessThan = sheet.DataValidations.AddTimeValidation("A1");
            lessThan.Operator = ExcelDataValidationOperator.equal;

            lessThan.Formula.Value.Hour = 10;

            ExcelTime time = new ExcelTime();
            time.Hour = 9;

            sheet.Cells["B1"].Value = time.ToExcelString();
            sheet.Cells["B1"].Style.Numberformat.Format = "HH:MM:SS";

            string timeString = lessThan.Formula.Value.ToExcelString();

            lessThan.Formula.ExcelFormula = $"=IF(B1<\"{timeString}\"," +
                $"\"{time.ToExcelString()}\",\"{timeString}\")";
            lessThan.ShowErrorMessage = true;

            var greaterThan = sheet.DataValidations.AddTimeValidation("A2");

            greaterThan.Formula.ExcelFormula = $"=B1>\"{time.ToExcelString()}\"";
            greaterThan.ShowErrorMessage = true;

            greaterThan.Operator = ExcelDataValidationOperator.equal;

            MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);

            var loadedpkg = new ExcelPackage(stream);
            var loadedSheet = loadedpkg.Workbook.Worksheets[0];

            var validations = loadedSheet.DataValidations;

            Assert.AreEqual(((ExcelDataValidationTime)validations[0]).Formula.ExcelFormula, 
                $"=IF(B1<\"{timeString}\"," +
                $"\"{time.ToExcelString()}\",\"{timeString}\")");
            Assert.AreEqual(((ExcelDataValidationTime)validations[1]).Formula.ExcelFormula, $"=B1>\"{time.ToExcelString()}\"");
        }
    }
}
