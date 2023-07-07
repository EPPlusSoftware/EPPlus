using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class AmanaIssues : TestBase
    {
        [TestMethod]
        public void ExcelPackage_SaveAs_doesnt_throw_exception()
        {
            // Arrange
            var input = GetTestStream("SN_T_1506944663_AufrissGewinnundVerlustrechnung.xlsx");
            var package = new ExcelPackage(input);
            var output = Path.GetTempFileName();

            // Act-Assert
            package.SaveAs(output);

            // Cleanup
            File.Delete(output);
        }

        [TestMethod]
        public void ExcelPackage_Calculate()
        {
            // Arrange
            var input = GetTestStream("Trim.xlsx");
            var package = new ExcelPackage(input);
            var sheet = package.Workbook.Worksheets[1];

            // Act
            sheet.Calculate();

            // Arrange
            Assert.AreEqual("Anlagevermögen", sheet.Cells["B8"].Value);
            Assert.AreEqual("123 456 ABC", sheet.Cells["B9"].Value);
        }
    }
}