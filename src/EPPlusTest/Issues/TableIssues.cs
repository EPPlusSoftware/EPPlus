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
        [TestMethod]
        public void i1314()
        {
            using (var p = OpenTemplatePackage("i1314.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var tbl = ws.Tables[0];
                tbl.InsertRow(1,1);
				tbl.AddRow(1);

				SaveAndCleanup(p);
            }
		}
    }
}
