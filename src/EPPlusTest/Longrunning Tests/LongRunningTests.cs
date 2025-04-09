using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace EPPlusTest.LongRunning
{

    [TestClass, Ignore]
    public class LongRunningTests : TestBase
    {

        const string SubFolder = "DigSig\\";

        X509Certificate2 GetSelfCert()
        {
            var requestedCert = new CertificateRequest("cn=SelfSignCert", RSA.Create(), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var finalCert = requestedCert.CreateSelfSigned(DateTimeOffset.Now, DateTimeOffset.Now.AddMinutes(5));

            var certPrivate = finalCert.Export(X509ContentType.Pfx);
            var certPublic = finalCert.Export(X509ContentType.Cert);
            var newCert = new X509Certificate2(certPrivate, "", X509KeyStorageFlags.Exportable);
            return newCert;
        }


        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestInitialize]
        public void Initialize()
        {
        }
        #region Digital signature Tests
        [TestMethod]
        public void SignSaveFileWithLOTSOfData()
        {
            string fileName = $"s350.xlsm";

            using (var pck = OpenTemplatePackage(fileName))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitalSignatures.Add(cert);

                pck.SaveAs(GetOutputFile(SubFolder, fileName));
            }
        }
        [TestMethod]
        public void SignAndVerifySigningInformation()
        {
            var title = "A Title";
            var address = "Some";
            var address2 = "Where";
            var ZIPorPostalCode = "Over";
            var city = "The";
            var CountryOrRegion = "Rainbow";
            var StateOrProvince = "WayUpHigh";

            string fileName = $"combineddatareport.xlsx";

            using (var pck = OpenTemplatePackage(fileName))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitalSignatures.Add(cert);
                var info = digSig.Details;

                info.SignerRoleTitle = title;
                info.Address1 = address;
                info.Address2 = address2;
                info.ZipOrPostalCode = ZIPorPostalCode;
                info.City = city;
                info.CountryOrRegion = CountryOrRegion;
                info.StateOrProvince = StateOrProvince;

                pck.SaveAs(GetOutputFile(SubFolder, fileName));
            }
            using (var pck = OpenPackage($"{SubFolder}{fileName}"))
            {
                var wb = pck.Workbook;
                var signerInformation = wb.DigitalSignatures[0].Details;
                Assert.AreEqual(title, signerInformation.SignerRoleTitle);
                Assert.AreEqual(address, signerInformation.Address1);
                Assert.AreEqual(address2, signerInformation.Address2);
                Assert.AreEqual(ZIPorPostalCode, signerInformation.ZipOrPostalCode);
                Assert.AreEqual(city, signerInformation.City);
                Assert.AreEqual(CountryOrRegion, signerInformation.CountryOrRegion);
                Assert.AreEqual(StateOrProvince, signerInformation.StateOrProvince);
            }
        }

        //Interestingly enough. Excel gets invalid signature when EXCEL tries to save this.
        //We do too
        [TestMethod]
        public void SignSaveFileWithLOTSOfData2()
        {
            using (var pck = OpenTemplatePackage("S610.xlsx"))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var digSig = wb.DigitalSignatures.Add(store.Certificates[1]);

                SaveAndCleanup(pck);
            }
        }
        #endregion
        #region Performance Tests
        [TestMethod]
        public void TableAddColumnToMax()
        {
            using (var p = new ExcelPackage()) // We discard this as it takes to long time to save
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("TableMaxColumn");
                LoadTestdata(ws, 100);
                var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableMaxColumn");
                //Act
                tbl.Columns.Add(ExcelPackage.MaxColumns - 4);
                //Assert
                Assert.AreEqual(ExcelPackage.MaxColumns, tbl.Address._toCol);
            }
        }
        [TestMethod]
        public void PerformanceIssueGetAsByteArray()
        {
            using (var p = OpenTemplatePackage("TemplateWithPivot.xlsx"))
            {
                /* Raw Data Sheet only */
                ExcelWorksheet ws = p.Workbook.Worksheets[1];  // second sheet

                // write data
                var table = ws.Tables[0];
                table.InsertRow(position: 1, rows: 6620);  // necessary to have the formulas available.

                // write data to buffer. This takes too long.
                var pt = p.Workbook.Worksheets[0].PivotTables[0];
                p.Workbook.Calculate();
                SaveWorkbook("PivotTest_calculated_columns.xlsx", p);
            }
        }
        #endregion
    }
}
