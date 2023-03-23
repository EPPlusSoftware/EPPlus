using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System.IO;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ExtLstValidationTests : TestBase
    {
        [TestMethod, Ignore]
        public void AddValidationWithFormulaOnOtherWorksheetShouldReturnExt()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                var sheet2 = package.Workbook.Worksheets.Add("test2");
                var val = sheet1.DataValidations.AddListValidation("A1");
                val.Formula.ExcelFormula = "test2!A1:A2";
                Assert.IsInstanceOfType(val, typeof(ExcelDataValidationList));
            }
        }

        [TestMethod, Ignore]
        public void ReadAndSaveExtLstPackage_ShouldNotThrow()
        {
            using (ExcelPackage package = OpenTemplatePackage("ExtLstDataValidationValidation.xlsx"))
            {
                SaveAndCleanup(package);

                var memoryStream = new MemoryStream();
                package.SaveAs(memoryStream);
                ExcelPackage p = new ExcelPackage(memoryStream);

                Assert.IsTrue(p.Workbook.Worksheets[0].DataValidations.Count > 0);
            }
        }
    }
}
