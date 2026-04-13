using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;
using System.Globalization;
using System.Diagnostics;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class RangeTextIssues : TestBase
    {
        [TestMethod]
        public void s667()
        {
            using (ExcelPackage package = OpenTemplatePackage("s667.xlsx"))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                /*
                 123+456	10-10-10	50,00 
                456+789	11-11-11	25,00 
                234+567	9-9-09	10,00 

                 */
                SwitchToCulture();
                Assert.AreEqual("123+456", worksheet.Cells["A2"].Text);
                Assert.AreEqual("456+789", worksheet.Cells["A3"].Text);
                Assert.AreEqual("234+567", worksheet.Cells["A4"].Text);

                Assert.AreEqual("10/10/10", worksheet.Cells["B2"].Text);
                Assert.AreEqual("11/11/11", worksheet.Cells["B3"].Text);
                Assert.AreEqual("9/9/09", worksheet.Cells["B4"].Text);

                Assert.AreEqual("50.00", worksheet.Cells["C2"].Text);
                Assert.AreEqual("25.00", worksheet.Cells["C3"].Text);
                Assert.AreEqual("10.00", worksheet.Cells["C4"].Text);

                SwitchBackToCurrentCulture();

                package.Save();

            }
        }
        [TestMethod]
        public void i1653()
        {
            using(var package = OpenTemplatePackage("i1653.xlsx"))
            {
                ExcelWorkbook workbook = package.Workbook;
                ExcelWorksheet worksheet = workbook.Worksheets[0];
                var cell = worksheet.Cells["C17"].Text;

                // Assert failed: `cell.Text` and `cell.Value` are both empty string
                    Debug.Assert(cell == "{\"any_orange_frag\" : 700, \"any_purple_frag\" : 800,\"any_blue_frag\" : 1200, \"any_green_frag\":2500}");
            }
        }
        [TestMethod]
        public void i1964()
        {
            using (var package = OpenTemplatePackage("WrongText-copy.xlsx"))
            {
                ExcelWorkbook workbook = package.Workbook;
                ExcelWorksheet worksheet = workbook.Worksheets[0];
                var cell = worksheet.Cells["A1"].Text;
                Assert.AreEqual("2021-12-31", cell);
            }
        }
        [TestMethod]
        public void i2276()
        {
            using (var package = OpenTemplatePackage("CurrencyTest2.xlsx"))
            {
                SwitchToCulture("nl-NL");
                ExcelWorkbook workbook = package.Workbook;
                ExcelWorksheet worksheet = workbook.Worksheets[0];

                Assert.AreEqual("€ 12 347", worksheet.Cells["A1"].Text);
                Assert.AreEqual("[$€-2]\\ #,##0", worksheet.Cells["A1"].Style.Numberformat.Format);
                Assert.AreEqual("€ 12 346,78", worksheet.Cells["D1"].Text);
                Assert.AreEqual("[$€-2]\\ #,##0.00", worksheet.Cells["D1"].Style.Numberformat.Format);

                Assert.AreEqual("€ 12.347", worksheet.Cells["A2"].Text);
                Assert.AreEqual("\"€\"\\ #,##0", worksheet.Cells["A2"].Style.Numberformat.Format);
                Assert.AreEqual("€ 12.346,78", worksheet.Cells["D2"].Text);
                Assert.AreEqual("\"€\"\\ #,##0.00", worksheet.Cells["D2"].Style.Numberformat.Format);
                SwitchBackToCurrentCulture();
            }
        }
    }
}
