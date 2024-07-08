using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class PivotTableIssues : TestBase
    {
        [TestMethod]
        public void s688()
        {
            using (ExcelPackage package = OpenTemplatePackage("s688.xlsx"))
            {
                package.Workbook.Worksheets[0].PivotTables[0].Calculate(true);
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s692()
        {
            using (ExcelPackage p = OpenTemplatePackage("s692.xlsx"))
            {
                foreach (ExcelWorksheet worksheet in p.Workbook.Worksheets)
                {
                    foreach (var table in worksheet.PivotTables)
                    {
                        table.Calculate(refreshCache: true);
                    }
                }
                SaveAndCleanup(p);
            }
        }
    }
}
