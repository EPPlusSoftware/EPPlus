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
            ExcelPackage.License.SetCommercial(lk);
            using(var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
            }
        }
        [TestMethod]
        public void CommercialFunction2Test()
        {
            var lk = "hnh8pj+e4dKUVwUwL2lW3b+4sP00YAF2lrE6W8BdD48HUTVGN3htPE8kdcIm+TEmwYm9YtBBcIbAQuJLIyl1+AEGQjM1OTc45QcfAYwCCgEA";
            ExcelPackage.License.SetCommercial(lk);
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
            }
        }
        [TestMethod]
        public void CommercialTrialSubsciptionTest()
        {
            var lk = "fX47BAtakq4T6v/K/zosjWipM9npn2yVWLhFn8MAsdDGJc2fN5+Lsd6rcRc4c1PzlF1IVX1UoDQbEkM+IahCAAEGQjA4RDc35QdWAMMBEQoA";
            ExcelPackage.License.SetCommercial(lk);
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                Assert.AreEqual(EPPlusCommercialLicenseType.Trial | EPPlusCommercialLicenseType.Subscription, ExcelPackage.License.LicenseInfo.LicenseType);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
            }
        }
        [TestMethod]
        public void CommercialSubscriptionExpired()
        {
            var lk = "ayWztKDgygaybN77mot4cA0NlE/QBf3riVa/OxuNFm6SbkkbJ1j3KJZyVRq3euJxg2LVbpZKlrc8rTgAgeM4wwEGQjhEQjAx6AdQAIIAAQIA";
            ExcelPackage.License.SetCommercial(lk);
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
            }
        }
        [TestMethod]
        public void CommercialTemporaryKey()
        {
            var lk = "tKChnon7eEepmLgXpVt0EhRO5dd/sDfvxLiZ+M5exmU3SjZh7Jj/Q8SHl59GUJoz0TL0xxS8IR7kfy1rD2N3FQEGQjFGQTE36AdKAbcCAQIA";
            ExcelPackage.License.SetCommercial(lk);
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
            }
        }        
        
        [TestMethod]
        public void CommercialTrialExpired()
        {
            var lk = "EIAhDzH0DqT+b827sKZKjnvz8dC3/4tu5tCr8/BeYoC6aMgR/0yIhTYBqZXg1sZbH60L1qZtvI39r3z9dkQAzQEGQjJDNzFG5wc8AaoCAQEA";
            ExcelPackage.License.SetCommercial(lk);
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
            }
        }        
        [TestMethod]
        public void NonCommercialOrganizationFunctionTest()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus.Org");
            using (var p = new ExcelPackage())
            {

                var ws = p.Workbook.Worksheets.Add("Sheet1");
                SaveWorkbook("LicenseKeyNonComercialOrg.xlsx", p);
            }
        }
        [TestMethod]
        public void NonCommercialPersonalFunctionTest()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Jan Källman");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                SaveWorkbook("LicenseKeyNonComercialPersonal.xlsx", p);
            }
        }
        [TestMethod]
        public void CommercialConfigFileTest()
        {
            ExcelPackage.License.RemoveActiveLicense();
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual("CGNCoSa1GgSHYvcsjVTU1W3ege0vwtl/9gFYj7qsBXsuVj9iqIHa9Deej4N/ZHnSkpNySdq7AQP0hCnfuTiMVQAGQjAxMTYw6AcAAG4BAQIA",ExcelPackage.License.LicenseKey);
                Assert.AreEqual(EPPlusLicenseSource.ConfigFile, ExcelPackage.License.Source);
                Assert.AreEqual(EPPlusLicenseType.Commercial, ExcelPackage.License.LicenseType);
                SaveWorkbook("LicenseKeyComercialConfig.xlsx", p);
            }
        }
        [TestMethod]
        public void CommercialConfigEnvironmentTest()
        {
            ExcelPackage.License.RemoveActiveLicense();
            Environment.SetEnvironmentVariable("EPPlusLicense", "NonCommercialPersonal:Jan Källman", EnvironmentVariableTarget.Process);
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(EPPlusLicenseType.NonCommercialPersonal, ExcelPackage.License.LicenseType);
                Assert.AreEqual(EPPlusLicenseSource.EnvironmentVariable, ExcelPackage.License.Source);
                Assert.AreEqual("Jan Källman", ExcelPackage.License.LegalName);
                SaveWorkbook("LicenseKeyEnvironment.xlsx", p);
            }
            Environment.SetEnvironmentVariable("EPPlusLicense", null);
        }
    }
}
