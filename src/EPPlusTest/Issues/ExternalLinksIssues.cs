using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
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
        [TestMethod]
        public void s2592()
        {
            using var lp = OpenPackage("I2592-linkto.xlsx", true);
            lp.Workbook.Worksheets.Add("Prices");
            SaveAndCleanup(lp, false);

            // other.xlsx is any workbook with a sheet called "Prices"
            using var package = OpenPackage("I2592.xlsx", true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Qty";
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 2;
            sheet.Tables.Add(sheet.Cells["A1:A3"], "Table1");
            package.Workbook.ExternalLinks.AddExternalWorkbook(new FileInfo(lp.File.FullName));
            sheet.Cells["C1"].Formula = "IF(A2 = 1, \"a\", 0)";
            sheet.Cells["C2"].Formula = "SUM(Table1[Qty])";
            sheet.Cells["C3"].Formula = "[1]Prices!$A$1+1";
            sheet.Calculate();
            using var reopened = new ExcelPackage(new MemoryStream(package.GetAsByteArray()));
            reopened.Workbook.ExternalLinks.RemoveAt(0);

            var ws = reopened.Workbook.Worksheets["Sheet1"];
            Assert.AreEqual(sheet.Cells["C1"].Formula, ws.Cells["C1"].Formula);
            Assert.AreEqual(sheet.Cells["C2"].Formula, ws.Cells["C2"].Formula);
            Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["C3"].Formula));
            Assert.AreEqual(1D, ws.Cells["C3"].Value);
        }
    }
}
