using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class TableIssues : TestBase
    {
        [TestMethod]
        public void s594()
        {
            using (ExcelPackage package = OpenTemplatePackage("s594.xlsx"))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["dg"];

                ExcelCalculationOption excelCalculationOption = new ExcelCalculationOption();
                excelCalculationOption.AllowCircularReferences = true;
                worksheet.Calculate(excelCalculationOption);

                Assert.AreNotEqual(0, worksheet.Cells["A1"].Text);

                package.Save();
            }
        }
    }

}
