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
            var lk = "X/NPJJG0WcsBTa6ZiCq19LyCk4wCqVlYH3cTtmQ5KSykt+rlceusxdN1p6uvDGP7rvPPNs0CGsjrNTaJEUR5LwEGQjhEQjAx6AdQAIIAAQIA";
            try
            {
                ExcelPackage.License.SetCommercial(lk);
            }
            catch(LicenseNotValidException)
            {
                Assert.AreEqual(EPPlusLicenseStatus.Expired, ExcelPackage.License.LicenseInfo.Status);
                Assert.AreEqual(EPPlusLicenseSource.Code, ExcelPackage.License.Source);
                Assert.AreEqual(new DateTime(2024, 3, 21).ToOADate(), ExcelPackage.License.LicenseInfo.LicenseValidFrom.ToOADate());
                Assert.AreEqual(new DateTime(2024, 6, 9).ToOADate(), ExcelPackage.License.LicenseInfo.LicenseValidTo.ToOADate());
                Assert.AreEqual(2, ExcelPackage.License.LicenseInfo.NumberOfLicensedDevelopers);
            }
            Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
        }
        [TestMethod]
        public void CommercialFunctionInvalidKeyTest()
        {
            var lk = "hnh8pj+e4dKUVwUwL2lW3b+4sP00YAF2lrE6W8BdD48HUTVGN3htPE8kdcIm+TEmwYm9YtBBcIbAQuJLIyl1+AEGQjM1OTc45QcfAYwCCgEA";
            Assert.ThrowsExactly<InvalidLicenseKeyException>(() =>
            {
                ExcelPackage.License.SetCommercial(lk);
                using (var p = new ExcelPackage())
                {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                SaveWorkbook("LicenseKeyComercial.xlsx", p);
                }
            });
        }

        [TestMethod]
        public void CommercialSubscriptionExpired()
        {
            var lk = "GRkMe8goQZjmGTsyxjTiLv4FSrwd+Sb1DO8KEJttMkOoutyzpZ+qteojECrmj+w4OLcUVtYvHd0GFC5z0KKkjwEGQjAzNDJF5Qd8AHIAAQIA";
            try
            {
                ExcelPackage.License.SetCommercial(lk);
            }
            catch (Exception ex)
            {
                Assert.IsInstanceOfType(ex, typeof(LicenseNotValidException));
            }
        }
        [TestMethod]
        public void CommercialTemporaryKey()
        {
            var lk = "R+WEJREh+kZCPuQDxJDNB96heFdV0hZLG6xvWYAeEfAZjBU5JXFoT2/+or9Uxwf5aDyJxGi+VgDw92x+Cr6sMgEGQjZCQjM15wfJAAADQQIA";
            Assert.ThrowsExactly<LicenseNotValidException>(() =>
            {
                ExcelPackage.License.SetCommercial(lk);
                using (var p = new ExcelPackage())
                {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(EPPlusCommercialLicenseType.TemporaryKey | EPPlusCommercialLicenseType.Subscription, ExcelPackage.License.LicenseInfo.LicenseType);
                Assert.AreEqual(lk, ExcelPackage.License.LicenseKey);
                }
            });
        }        
        
        [TestMethod]
        public void CommercialTrialExpiredShouldReadFromConfig()
        {
            var lk = "yYQQXbIcxblb4BYxKfE3wPRvnb+r7ZEZZaTgDHE2QfHDpcsNb4dJJT3lz56I/MbjcCfJ9d8aG5+teLoQVIAx4gEGQjM5MTg45gcwAVEBIQEA";
            ExcelPackage.License.RemoveActiveLicense();
            try
            {
                ExcelPackage.License.SetCommercial(lk);
            }
            catch 
            {
                Assert.AreEqual(EPPlusLicenseStatus.Expired, ExcelPackage.License.LicenseInfo.Status);
                Assert.AreEqual(EPPlusLicenseSource.Code, ExcelPackage.License.Source);
                Assert.AreEqual(EPPlusCommercialLicenseType.Trial | EPPlusCommercialLicenseType.Subscription, ExcelPackage.License.LicenseInfo.LicenseType);
            }
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                Assert.AreEqual(EPPlusLicenseSource.ConfigFile, ExcelPackage.License.Source);
                Assert.IsNull(ExcelPackage.License.LicenseKey);
                SaveAndCleanup(p);
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
                Assert.AreEqual("EPPlus Test Project", ExcelPackage.License.LegalName);
                Assert.AreEqual(EPPlusLicenseSource.ConfigFile, ExcelPackage.License.Source);
                Assert.AreEqual(EPPlusLicenseType.NonCommercialPersonal, ExcelPackage.License.LicenseType);
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
