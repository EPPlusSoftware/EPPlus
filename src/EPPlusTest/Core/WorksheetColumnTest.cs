using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
namespace EPPlusTest.Core
{
    [TestClass]
    public class WorksheetColumnTest : TestBase
    {
        [TestMethod]
        public void ValidateDefaultWidth()
        {
            using(var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(9.140625D, ws.DefaultColWidth);
            }
        }
    }
}
