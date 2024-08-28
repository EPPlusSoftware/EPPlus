using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class EncryptionTests : TestBase
    {
        [TestMethod]
        public void SensitivityLableRead()
        {
            var fi = GetTemplateFile("SensitivityLabel.xlsx");
            using(var p=new ExcelPackage(fi,""))
            {
                Assert.AreEqual(1, p.Workbook.Worksheets.Count);
            }
        }
    }
}
