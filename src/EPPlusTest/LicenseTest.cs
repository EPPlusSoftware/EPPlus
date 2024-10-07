using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class LicenseTest : TestBase
    {
        [TestMethod]
        public void CommercialFunctionTest()
        {
            var lk = "FRy3bIoLtKBhSmohLRw04TUBOkjldZpZ2njfJx3c9b/85NcTs1TT7Up6RCDEUSf9+lgv9KMLgABTOBBL/YY0FAAGQjAxMTZG6AcAAG4BAQUA";
            ExcelPackage.License.SetLicenseCommercial(lk);
            using(var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
            }
        }
        [TestMethod]
        public void NonCommercialOrganizationFunctionTest()
        {
            ExcelPackage.License.SetLicenseNonCommercialOrganization("EPPlus.Org");
            using (var p = new ExcelPackage())
            {

                var ws = p.Workbook.Worksheets.Add("Sheet1");
                SaveWorkbook("LicenseKeyNonComercialOrg.xlsx", p);
            }
        }
        [TestMethod]
        public void NonCommercialPersonalFunctionTest()
        {
            ExcelPackage.License.SetLicenseNonCommercialPersonal("Jan Källman");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                SaveWorkbook("LicenseKeyNonComercialPersonal.xlsx", p);
            }
        }
        [TestMethod]
        public void CommercialConfigFileTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual("CGNCoSa1GgSHYvcsjVTU1W3ege0vwtl/9gFYj7qsBXsuVj9iqIHa9Deej4N/ZHnSkpNySdq7AQP0hCnfuTiMVQAGQjAxMTYw6AcAAG4BAQIA",ExcelPackage.License.LicenseKey);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
            }
        }
    }
}
