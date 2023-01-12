using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml.DataValidation;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class Test635 : TestBase
    {
        [TestMethod]
        public void Stuff()
        {
            using (var P = OpenPackage("blahFnutt.xlsx", false))
            {
                ExcelWorksheet worksheet = P.Workbook.Worksheets[1];
                //P.Workbook.Worksheets.Add("NewSheet");
                //AddIntegerValidation(P);

                //SaveAndCleanup(P);
            }

        }

        private static void AddIntegerValidation(ExcelPackage package) 
        {
            var sheet = package.Workbook.Worksheets.Add("integer");
            //add a validation and set values
            var validation = sheet.DataValidations.AddIntegerValidation("A1:A2");

            // Alternatively:
            //var validation = sheet.Cells["A1:A2"].DataValidation;

            //validation.AddAnyDataValidation();
            //var validation2 = validation.AddDateTimeDataValidation();

            //validation2.ErrorStyle = ExcelDataValidationWarningStyle.stop;


            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.PromptTitle = "Enter a integer value here";
            validation.Prompt = "Value should be between 1 and 5";
            validation.ShowInputMessage = true;
            validation.ErrorTitle = "An invalid value was entered";
            validation.Error = "Value must be between 1 and 5";
            validation.ShowErrorMessage = true;
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 5;
        }
    }
}
