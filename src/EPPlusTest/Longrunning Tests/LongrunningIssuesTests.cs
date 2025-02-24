using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlusTest.LongrunningTests
{
    [TestClass, Ignore]
    internal class LongrunningIssuesTests : TestBase
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
        [TestMethod]
        public void s350()
        {
            using (var p = OpenTemplatePackage("s350.xlsm"))
            {
                SaveWorkbook("s350.xlsm", p);
            }
        }
        [TestMethod]
        public void Issue294()
        {
            using (var p = OpenTemplatePackage("test_excel_workbook_before2-xl.xlsx"))
            {
                var s = p.Workbook.Styles.NamedStyles.Count;
                var ws = p.Workbook.Worksheets["Summary"];
                p.Save();
            }
        }
        [TestMethod]
        public void s551_2()
        {
            using (var p = OpenTemplatePackage("s551.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var usedRange = ws.Cells["a1:b5"];
                foreach (ExcelRangeRow dataRow in usedRange.EntireRow)
                {
                    if (dataRow.Hidden == false)
                    {
                        dataRow.Range.Formula = "f1";
                    }
                }
            }
        }
        [TestMethod]
        public void i863()
        {
            using (var p = OpenTemplatePackage("i863.xlsx"))
            {
                // Removed insertion of PHI data, just re-saving the template for sample purposes

                // Workaround - Issue with "Inputs" tab - Validation of T60:T64 failed: Formula2 must be set if operator is 'between' or 'notBetween' when cells are not using between or notBetween
                var otherInputTab = p.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals("Inputs"));
                if (otherInputTab != null)
                {
                    otherInputTab.DataValidations.InternalValidationEnabled = false;
                }
                // Saving
                SaveAndCleanup(p);

                var p2 = OpenPackage("i863.xlsx");

                var ws17 = p2.Workbook.Worksheets[16];
            }
        }
        [TestMethod]
        public void s539()
        {
            //Outputs
            var pc = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

                string sheetName = "Sheet1";
                string range = "G2:G5";
                string value = "VLOOKUP(F2,'Reference Data'!A2:B187021,2,0)";
                var logFile = new FileInfo("c:\\temp\\formulaLog.log");
                if (logFile.Exists) logFile.Delete();
                using (var package = OpenTemplatePackage("s539.xlsm"))
                {
                    package.Workbook.FormulaParserManager.AttachLogger(logFile);
                    var ws = package.Workbook.Worksheets[sheetName];
                    ws.Cells[range].Formula = value;
                    ws.Cells[range].Calculate();
                    SaveAndCleanup(package);
                }
            }
            catch (Exception e)
            {
                string exc = "";
                exc = "Failed. " + e.ToString();
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = pc;
                System.GC.Collect();
            }
        }
        [TestMethod]
        public void s610()
        {
            using (var p = OpenTemplatePackage("s610.xlsx"))
            {
                var wTestSheet = p.Workbook.Worksheets[0];
                //wTestSheet.Name = "Sheet2";
                //wTestSheet.View.UnFreezePanes();
                wTestSheet.InsertColumn(1, 2);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s614()
        {
            using (var package = OpenTemplatePackage("s614.xlsx"))
            {
                int sheetIndex = 5;
                var sheetName = $"Data Sheet_{sheetIndex}";
                var worksheet = package.Workbook.Worksheets[sheetName];
                worksheet.Name = "TestSheet_{sheetIndex}";

                worksheet.InsertColumn(1, 2);
                worksheet.Cells.Style.Font.Name = "ＭＳ Ｐゴシック";
                worksheet.Cells.Style.Font.Size = 11;

                worksheet.Cells[1, 1].Value = "TextTextTextTextTextTextTextTextTextTextTextText";

                worksheet.Column(1).AutoFit();
                worksheet.Column(2).AutoFit();

                package.Save();
            }
        }
        [TestMethod]
        public void s789()
        {
            using (var package = OpenTemplatePackage("s789.xlsx"))
            {
                var wb = package.Workbook;
                foreach (var ws in package.Workbook.Worksheets)
                {
                    foreach (var pTable in ws.PivotTables)
                    {
                        foreach (var field in pTable.Fields)
                        {
                        }
                    }
                }

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void SignSaveFileWithLOTSOfData()
        {
            string fileName = $"s350.xlsm";

            using (var pck = OpenTemplatePackage(fileName))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitialSignatures.Add(cert);

                pck.SaveAs(GetOutputFile(SubFolder, fileName));
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
                var digSig = wb.DigitialSignatures.Add(store.Certificates[1]);

                SaveAndCleanup(pck);
            }
        }
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
    }
}
