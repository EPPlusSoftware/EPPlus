using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class NegationIssues : TestBase
    {
        [TestMethod]
        public void s594()
        {
            string filePath = "C:\\Users\\OssianEdström\\Downloads\\dg-error.xlsx";//dg.xlsx is ok
            string sheetName = "dg";
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                ExcelCalculationOption excelCalculationOption = new ExcelCalculationOption();
                excelCalculationOption.AllowCircularReferences = true;
                worksheet.Calculate(excelCalculationOption);
                var text = worksheet.Cells["A1"].Text;

                package.Save();
            }
        }
    }
}
