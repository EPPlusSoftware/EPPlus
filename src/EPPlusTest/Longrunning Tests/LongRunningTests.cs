using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

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
        [TestMethod]
        public void PerformanceIssueLoadAndSave()
        {
            using (var p = OpenTemplatePackage("LargeWorkbookTemplate.xlsx"))
            {
                /* Raw Data Sheet only */
                ExcelWorksheet ws = p.Workbook.Worksheets[0];  // second sheet

                p.Workbook.Calculate();
                SaveWorkbook("LargeWBSave.xlsx", p);
            }
        }
        [TestMethod]
        public void PerformanceIssueLoadAndSaveSync()
        {
            using (var p = new ExcelPackage())
            {
                var file = GetTemplateFile("LargeWorkbookTemplate.xlsx");
                p.Load(new FileStream(file.FullName, FileMode.Open));
                /* Raw Data Sheet only */
                ExcelWorksheet ws = p.Workbook.Worksheets[0];  // second sheet

                p.Workbook.Calculate();
            }
        }


        [TestMethod]
        public async Task PerformanceIssueLoadAndSaveAsync()
        {
            using (var p = new ExcelPackage())
            {
                var file = GetTemplateFile("LargeWorkbookTemplate.xlsx");
                await p.LoadAsync(file);
                /* Raw Data Sheet only */
                ExcelWorksheet ws = p.Workbook.Worksheets[0];  // second sheet

                p.Workbook.Calculate();                
            }
        }

        #endregion
        #region HtmlExport
        [TestMethod]
        public async Task WriteAdvancedWs()
        {
            string _htmlOutput;
            _htmlOutput = _worksheetPath + "\\html\\";
            if (Directory.Exists(_htmlOutput) == false)
            {
                Directory.CreateDirectory(_htmlOutput);
            }

            using (var p = OpenTemplatePackage("s610.xlsx"))
            {
                var sheet1 = p.Workbook.Worksheets[0];
                var exporterRange = p.Workbook.CreateHtmlExporter(sheet1.Cells["A1:BL7868"]);
                exporterRange.Settings.SetColumnWidth = true;
                exporterRange.Settings.SetRowHeight = true;
                exporterRange.Settings.Minify = false;
                exporterRange.Settings.TableStyle = eHtmlRangeTableInclude.Include;
                exporterRange.Settings.Pictures.Include = ePictureInclude.Include;
                var htmlAsync = await exporterRange.GetSinglePageAsync();

                File.WriteAllText($"{_htmlOutput}RangeAndThreeTables.html", htmlAsync);
            }
        }
        #endregion
        #region PivotCacheStoreTests
        [TestMethod]
        public void s789_charts()
        {
            using (var pck = OpenTemplatePackage("s789_charts_justOne.xlsx"))
            {
                var wb = pck.Workbook;
                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void s789_orignals()
        {
            using (var pck = OpenTemplatePackage("s789_original_issue.xlsx"))
            {
                var wb = pck.Workbook;

                var ws = wb.Worksheets.GetByName("PivotTables");
                var table = ws.PivotTables[0];

                SaveAndCleanup(pck);
            }
        }
        #endregion
        #region chartInsertTests
        [TestMethod]
        public void ColumnCheck()
        {
            using (var p = OpenTemplatePackage("s808_2.xlsx"))
            {
                var ws = p.Workbook.Worksheets["overzicht"];

                ws.Calculate();

                ws.ClearFormulas();

                ws.Columns.DeleteAll(c => c.Hidden);

                SaveAndCleanup(p);
            }
        }
        #endregion
    }
}
