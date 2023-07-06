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
        public void Calculate_calculates_formula_with_external_link()
        {
            // Arrange
            var input = GetTestStream("ExternalReferences.xlsx");
            var package = new ExcelPackage(input);
            var sheet = package.Workbook.Worksheets[0];

            // Act
            sheet.Calculate();

            // Assert
            Assert.AreEqual(60d, sheet.Cells["A1"].Value);
            Assert.AreEqual(60d, sheet.Cells["A2"].Value);
            Assert.AreEqual(23d, sheet.Cells["B19"].Value);
            Assert.AreEqual(23d, sheet.Cells["B20"].Value);
        }
    }
}