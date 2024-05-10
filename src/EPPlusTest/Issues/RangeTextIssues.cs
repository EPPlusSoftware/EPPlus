using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;
using System.Globalization;

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
 
    }
}
