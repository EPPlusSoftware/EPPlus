using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
namespace EPPlusTest.Issues
{
    [TestClass]
    public class ExternalLinksIssues : TestBase
    {
        [TestMethod, Ignore]
        public void s786()
        {
            using (var package = OpenTemplatePackage("s786.xlsx"))
            {
                FileInfo externalLinkFile = new FileInfo("D:\\test.xlsx");
                if (externalLinkFile != null)
                {
                    var externalWorkbook = package.Workbook.ExternalLinks.AddExternalWorkbook(externalLinkFile);

                }
                var worksheet = package.Workbook.Worksheets["Sheet1"];
                worksheet.Cells["D1000"].Formula = "[9]Sheet!C2";
                SaveAndCleanup(package);
            }
            using (ExcelPackage package = OpenPackage("s786.xlsx"))
            {
                foreach (var item in package.Workbook.ExternalLinks)
                {
                    if (item.As.ExternalWorkbook.File == null)
                    {
                        continue;
                    }
                    string pathFile = item.As.ExternalWorkbook.File.FullName;
                }
            }
        }
    }
}
