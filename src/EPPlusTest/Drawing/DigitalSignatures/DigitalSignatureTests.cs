using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security.Cryptography;
using System.Xml;
using System.IO;
using System;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Text;
using OfficeOpenXml.Utils;

//REMEMBER:
//1. Cannonize
//2. Transform
//3. Hash data

//TODO: Write tests to check that individual reference hashes of the signature are correctly generated internally
//One for package, office etc.

namespace EPPlusTest.Drawing.DigitalSignatures
{
    [TestClass]
    public class DigitalSignatureTests : TestBase
    {
        X509Certificate2 GetSelfCert()
        {
            X509Store store = GetStore();

            if(store.Certificates.Count == 0)
            {
                var requestedCert = new CertificateRequest("cn=SelfSignCert", RSA.Create(), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                var finalCert = requestedCert.CreateSelfSigned(DateTimeOffset.Now, DateTimeOffset.Now.AddMinutes(5));

                var certPrivate = finalCert.Export(X509ContentType.Pfx);
                var newCert = new X509Certificate2(certPrivate, "", X509KeyStorageFlags.Exportable);

                store.Add(newCert);
            }
            return store.Certificates[0];
        }

        static X509Store GetStore()
        {
            X509Store store = new X509Store("tmpStoreDigSigEpplus", StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadWrite);
            return store;
        }

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            X509Store store = GetStore();

            foreach (var cert in store.Certificates)
            {
                store.Remove(cert);
            }
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            X509Store store = GetStore();
            foreach (var cert in store.Certificates)
            {
                store.Remove(cert);
            }
            store.Close();
        }

        [TestMethod]
        public void CreateDigitalSignatureAndReadIt()
        {
            using (ExcelPackage package = OpenPackage("InvisibleSignature.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                wb.Calculate();

                var test = package.Workbook.FullCalcOnLoad;

                var cert = GetSelfCert();

                var digSig = ws.Workbook.DigitialSignatures.AddSignature(cert, CommitmentType.CreatedAndApproved, "TestingSignatureLine");
                var info = digSig.Details;

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZIPorPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                SaveAndCleanup(package);
                Assert.IsTrue(digSig.IsValid);
            }
            using (ExcelPackage package = OpenPackage("InvisibleSignature.xlsx"))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets[0];

                var cert = GetSelfCert();

                var digSig = wb.DigitialSignatures[0];
                var info = digSig.Details;
                Assert.AreEqual("A Title", info.SignerRoleTitle);
                Assert.AreEqual("Some", info.Address1);
                Assert.AreEqual("Where", info.Address2);
                Assert.AreEqual("Over", info.ZIPorPostalCode);
                Assert.AreEqual("The", info.City);
                Assert.AreEqual("Rainbow", info.CountryOrRegion);
                Assert.AreEqual("WayUpHigh", info.StateOrProvince);
                Assert.AreEqual(cert, digSig.Certificate);

                SaveAndCleanup(package);
                Assert.IsTrue(digSig.IsValid);
            }
        }

        [TestMethod]
        public void ReadCommitmentTypeAndTypeQualifier()
        {
            using (ExcelPackage package = OpenTemplatePackage("DigSig_FullSignatureAndLine.xlsx"))
            {
                var wb = package.Workbook;

                var digSig = wb.DigitialSignatures[0];
                Assert.AreEqual(CommitmentType.Approved, digSig.CommitmentTyping);
                Assert.AreEqual("MyPurposeIsMyOwn", digSig.PurposeForSigning);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void ReadCommitmentTypeAndTypeQualifierWhenNone()
        {
            using (ExcelPackage package = OpenTemplatePackage("DigSig_FullSignatureAndLineNone.xlsx"))
            {
                var wb = package.Workbook;

                var digSig = wb.DigitialSignatures[0];
                Assert.AreEqual(CommitmentType.None, digSig.CommitmentTyping);
                Assert.AreEqual("MyPurposeIsMyOwn", digSig.PurposeForSigning);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void EnsureSignatureReferencesAreEncodedCorrectly()
        {
            using (ExcelPackage package = OpenPackage("NewOfficeReference.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("ANewWorksheet");

                //Add data, pivot table and chart so that Package Reference in digital signature can have more files to hash.
                //--------------------------------------------------BEGIN--------------------------------------------------------
                ws.Cells["A1"].Value = "PointsA";
                ws.Cells["B1"].Value = "PointsB";
                ws.Cells["C1"].Value = "PointsC";

                for (int i = 2; i <= 100; i++)
                {
                    for (int j = 1; j <= 100; j++)
                    {
                        ws.Cells[i, j].Value = i + j;
                    }
                }

                var pvWs = wb.Worksheets.Add("PivotTableWorksheet");

                var pt = pvWs.PivotTables.Add(pvWs.Cells["A1"], ws.Cells["A1:C10"], "APivotTable");

                pt.RowFields.Add(pt.Fields["PointsA"]);
                pt.DataFields.Add(pt.Fields["PointsB"]);
                pt.DataOnRows = true;

                var chart = pvWs.Drawings.AddPieChart("PivotChart", ePieChartType.PieExploded3D, pt);
                chart.SetPosition(1, 0, 4, 0);
                chart.SetSize(800, 600);
                chart.Legend.Remove();
                chart.Series[0].DataLabel.ShowCategory = true;
                chart.Series[0].DataLabel.Position = eLabelPosition.OutEnd;
                chart.StyleManager.SetChartStyle(ePresetChartStyle.Pie3dChartStyle6);
                //--------------------------------------------------End--------------------------------------------------------

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                wb.DigitialSignatures.AddSignature(store.Certificates[0], CommitmentType.CreatedAndApproved, "ToCompareDigitalSignatures");

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void EnsureEpplusHashesRelsTransformsCorrectly()
        {
            //.rels file
            string DotRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"_xmlsignatures/origin.sigs\"/></Relationships>";

            PartWithXml xml = new PartWithXml() { UriKey = "/_rels/.rels?ContentType=application/vnd.openxmlformats-package.relationships+xml", StringData = DotRels, PartType = ePartType.RelPart };

            var manifestReference = new ManifestReference(xml, DigitalSignatureHashAlgorithm.SHA1);

            var EpplusRelReference = manifestReference.xmlDigSig;

            string ExcelRelReference = "<Reference URI=\"/_rels/.rels?ContentType=application/vnd.openxmlformats-package.relationships+xml\"><Transforms><Transform  Algorithm=\"http://schemas.openxmlformats.org/package/2006/RelationshipTransform\"><mdssi:RelationshipReference xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\" SourceId=\"rId1\" /></Transform><Transform Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" /></Transforms><DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\" /><DigestValue>+nAd0bim5u961Z6hkrztwiSj8HA=</DigestValue></Reference>";
            var excelDoc = new XmlDocument();
            excelDoc.LoadXml(ExcelRelReference);

            Assert.AreEqual(excelDoc.InnerText, EpplusRelReference.InnerText);
        }

        [TestMethod]
        public void EnsureEpplusHashesDrawingsCorrectly()
        {
            //Drawing1 file
            string DrawingXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>16</xdr:col><xdr:colOff>304800</xdr:colOff><xdr:row>31</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"PivotChart\"><a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{00000000-0008-0000-0100-000002000000}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";
            var uriQuery = "/xl/drawings/_rels/drawing1.xml.rels?ContentType=application/vnd.openxmlformats-package.relationships+xml";

            PartWithXml xml = new PartWithXml() { UriKey = uriQuery, StringData = EncodeUtil.HashAndEncodeBytes(Encoding.UTF8.GetBytes(DrawingXml)), PartType = ePartType.Part };

            var manifestReference = new ManifestReference(xml, DigitalSignatureHashAlgorithm.SHA1);

            var EpplusRelReference = manifestReference.xmlDigSig;

            string ExcelRelReference = "<Reference URI=\"/xl/drawings/drawing1.xml?ContentType=application/vnd.openxmlformats-officedocument.drawing+xml\"> <DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\" /><DigestValue>9PPC/LpKQHYwAJRKNyzzxQfRZ3I=</DigestValue></Reference>";
            var excelDoc = new XmlDocument();
            excelDoc.LoadXml(ExcelRelReference);

            Assert.AreEqual(excelDoc.InnerText, EpplusRelReference.InnerText);
        }

        [TestMethod]
        public void EnsureSignatureReferencesAreEncodedCorrectly2()
        {
            using (ExcelPackage package = OpenTemplatePackage("ExcelFileToSign.xlsx"))
            {
                var wb = package.Workbook;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                wb.DigitialSignatures.AddSignature(store.Certificates[0], CommitmentType.CreatedAndApproved, "Compare");

                SaveAndCleanup(package);
            }

            //Open signed package
            using (ExcelPackage package = OpenPackage("ExcelFileToSign.xlsx"))
            {

            }
        }

        [TestMethod]
        public void CreateDigitalSignatureLineAndSignIt()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "ASigner";

                var digSig = ws.Workbook.DigitialSignatures.AddSignature(store.Certificates[1], CommitmentType.CreatedAndApproved, "TestingSignatureLine");
                var info = digSig.Details;

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZIPorPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                digSig.SignatureLine = sLine;

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void VerifyEncodingOfEmf()
        {
            var fullName = GetTemplateFile("InvalidImageOriginal.emf").FullName;
            var bytes = File.ReadAllBytes(fullName);
            var invalidImage = Convert.ToBase64String(bytes, Base64FormattingOptions.None);
            var originalInvalidImage = "AQAAAGwAAAAAAAAAAAAAAP8AAAB/AAAAAAAAAAAAAABcFwAAqwsAACBFTUYAAAEAsB8AALEAAAAGAAAAAAAAAAAAAAAAAAAAAAoAAKAFAABWAgAAUAEAAAAAAAAAAAAAAAAAAPAfCQCAIAUACgAAABAAAAAAAAAAAAAAAEsAAAAQAAAAAAAAAAUAAAAeAAAAGAAAAAAAAAAAAAAAAAEAAIAAAAAnAAAAGAAAAAEAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAEAAABMAAAAZAAAAAAAAAAAAAAA/wAAAH8AAAAAAAAAAAAAAAABAACAAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAABAAAAJwAAABgAAAABAAAAAAAAAP///wAAAAAAJQAAAAwAAAABAAAATAAAAGQAAAAAAAAAAAAAAP8AAAB/AAAAAAAAAAAAAAAAAQAAgAAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAQAAACcAAAAYAAAAAQAAAAAAAADw8PAAAAAAACUAAAAMAAAAAQAAAEwAAABkAAAAAAAAAAAAAAD/AAAAfwAAAAAAAAAAAAAAAAEAAIAAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAEAAAAnAAAAGAAAAAEAAAAAAAAA8PDwAAAAAAAlAAAADAAAAAEAAABMAAAAZAAAAAAAAAAAAAAA/wAAAH8AAAAAAAAAAAAAAAABAACAAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAABAAAAJwAAABgAAAABAAAAAAAAAPDw8AAAAAAAJQAAAAwAAAABAAAATAAAAGQAAAAAAAAAAAAAAP8AAAB/AAAAAAAAAAAAAAAAAQAAgAAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAQAAACcAAAAYAAAAAQAAAAAAAADw8PAAAAAAACUAAAAMAAAAAQAAAEwAAABkAAAAAAAAAAAAAAD/AAAAfwAAAAAAAAAAAAAAAAEAAIAAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAEAAAAnAAAAGAAAAAEAAAAAAAAA////AAAAAAAlAAAADAAAAAEAAABMAAAAZAAAAAAAAAAAAAAA/wAAAH8AAAAAAAAAAAAAAAABAACAAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAABAAAAJwAAABgAAAABAAAAAAAAAP///wAAAAAAJQAAAAwAAAABAAAATAAAAGQAAAAAAAAAAAAAAP8AAAB/AAAAAAAAAAAAAAAAAQAAgAAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAQAAACcAAAAYAAAAAQAAAAAAAAD///8AAAAAACUAAAAMAAAAAQAAAEwAAABkAAAAAAAAAAMAAAD/AAAAEgAAAAAAAAADAAAAAAEAABAAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAEAAAAnAAAAGAAAAAEAAAAAAAAA////AAAAAAAlAAAADAAAAAEAAABMAAAAZAAAAAkAAAADAAAAGAAAABIAAAAJAAAAAwAAABAAAAAQAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAABAAAAFQAAAAwAAAADAAAAcgAAALADAAAKAAAAAwAAABcAAAAQAAAA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AQICIlMAeQBzAHQAZQBtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZHYACAAAAAAlAAAADAAAAAIAAAAnAAAAGAAAAAMAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAMAAABMAAAAZAAAAAAAAAAAAAAA//////////8AAAAAFgAAAAAAAAA1AAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAADAAAAJwAAABgAAAADAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAADAAAATAAAAGQAAAAAAAAAAAAAAP//////////AAAAABYAAAAAAQAAAAAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAwAAACcAAAAYAAAAAwAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAwAAAEwAAABkAAAAAAAAAAAAAAD//////////wABAAAWAAAAAAAAADUAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAMAAAAnAAAAGAAAAAMAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAMAAABMAAAAZAAAAAAAAABLAAAA/wAAAEwAAAAAAAAASwAAAAABAAACAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAADAAAAJwAAABgAAAADAAAAAAAAAP///wAAAAAAJQAAAAwAAAADAAAATAAAAGQAAAAAAAAAFgAAAP8AAABKAAAAAAAAABYAAAAAAQAANQAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAwAAACcAAAAYAAAAAwAAAAAAAAD///8AAAAAACUAAAAMAAAAAwAAAEwAAABkAAAACQAAACcAAAAfAAAASgAAAAkAAAAnAAAAFwAAACQAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAMAAABSAAAAcAEAAAMAAADg////AAAAAAAAAAAAAAAAkAEAAAAAAAEAAAAAYQByAGkAYQBsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkdgAIAAAAACUAAAAMAAAAAwAAABgAAAAMAAAAAAAAABIAAAAMAAAAAQAAABYAAAAMAAAACAAAAFQAAABUAAAACgAAACcAAAAeAAAASgAAAAEAAAAA4LpBq6q6QQoAAABLAAAAAQAAAEwAAAAEAAAACQAAACcAAAAgAAAASwAAAFAAAABYAAAAFQAAABYAAAAMAAAAAAAAACUAAAAMAAAAAgAAACcAAAAYAAAABAAAAAAAAAD///8AAAAAACUAAAAMAAAABAAAAEwAAABkAAAAKQAAABkAAAD2AAAASgAAACkAAAAZAAAAzgAAADIAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAQAAAAnAAAAGAAAAAQAAAAAAAAA////AAAAAAAlAAAADAAAAAQAAABMAAAAZAAAACkAAAAZAAAA9gAAAEcAAAApAAAAGQAAAM4AAAAvAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAAEAAAAJwAAABgAAAAEAAAAAAAAAP///wAAAAAAJQAAAAwAAAAEAAAATAAAAGQAAAApAAAAMwAAAFkAAABHAAAAKQAAADMAAAAxAAAAFQAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAABAAAAFIAAABwAQAABAAAAPD///8AAAAAAAAAAAAAAACQAQAAAAAAAQAAAABzAGUAZwBvAGUAIAB1AGkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGR2AAgAAAAAJQAAAAwAAAAEAAAAGAAAAAwAAAAAAAAAEgAAAAwAAAABAAAAHgAAABgAAAApAAAAMwAAAFoAAABIAAAAJQAAAAwAAAAEAAAAVAAAAHAAAAAqAAAAMwAAAFgAAABHAAAAAQAAAADgukGrqrpBKgAAADMAAAAGAAAATAAAAAAAAAAAAAAAAAAAAP//////////WAAAAE8AcwBzAGkAYQBuAAwAAAAHAAAABwAAAAQAAAAIAAAACQAAAEsAAABAAAAAMAAAAAUAAAAgAAAAAQAAAAEAAAAQAAAAAAAAAAAAAAAAAQAAgAAAAAAAAAAAAAAAAAEAAIAAAAAlAAAADAAAAAIAAAAnAAAAGAAAAAUAAAAAAAAA////AAAAAAAlAAAADAAAAAUAAABMAAAAZAAAAAAAAABQAAAA/wAAAHwAAAAAAAAAUAAAAAABAAAtAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAAFAAAAJwAAABgAAAAFAAAAAAAAAP///wAAAAAAJQAAAAwAAAAFAAAATAAAAGQAAAAJAAAAUAAAAPYAAABcAAAACQAAAFAAAADuAAAADQAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAABQAAACUAAAAMAAAAAQAAABgAAAAMAAAAAAAAABIAAAAMAAAAAQAAAB4AAAAYAAAACQAAAFAAAAD3AAAAXQAAACUAAAAMAAAAAQAAAFQAAACoAAAACgAAAFAAAABhAAAAXAAAAAEAAAAA4LpBq6q6QQoAAABQAAAADwAAAEwAAAAAAAAAAAAAAAAAAAD//////////2wAAABTAHUAZwBnAGUAcwB0AGUAZABTAGkAZwBuAGUAcgAtQgYAAAAHAAAABwAAAAcAAAAGAAAABQAAAAQAAAAGAAAABwAAAAYAAAADAAAABwAAAAcAAAAGAAAABAAAAEsAAABAAAAAMAAAAAUAAAAgAAAAAQAAAAEAAAAQAAAAAAAAAAAAAAAAAQAAgAAAAAAAAAAAAAAAAAEAAIAAAAAlAAAADAAAAAIAAAAnAAAAGAAAAAUAAAAAAAAA////AAAAAAAlAAAADAAAAAUAAABMAAAAZAAAAAkAAABgAAAA9gAAAGwAAAAJAAAAYAAAAO4AAAANAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAAFAAAAJQAAAAwAAAABAAAAGAAAAAwAAAAAAAAAEgAAAAwAAAABAAAAHgAAABgAAAAJAAAAYAAAAPcAAABtAAAAJQAAAAwAAAABAAAAVAAAAKAAAAAKAAAAYAAAAFYAAABsAAAAAQAAAADgukGrqrpBCgAAAGAAAAAOAAAATAAAAAAAAAAAAAAAAAAAAP//////////aAAAAFMAdQBnAGcAZQBzAHQAZQBkAFQAaQB0AGwAZQAGAAAABwAAAAcAAAAHAAAABgAAAAUAAAAEAAAABgAAAAcAAAAGAAAAAwAAAAQAAAADAAAABgAAAEsAAABAAAAAMAAAAAUAAAAgAAAAAQAAAAEAAAAQAAAAAAAAAAAAAAAAAQAAgAAAAAAAAAAAAAAAAAEAAIAAAAAlAAAADAAAAAIAAAAnAAAAGAAAAAUAAAAAAAAA////AAAAAAAlAAAADAAAAAUAAABMAAAAZAAAAAkAAABwAAAAkAAAAHwAAAAJAAAAcAAAAIgAAAANAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAAFAAAAJQAAAAwAAAABAAAAGAAAAAwAAAAAAAAAEgAAAAwAAAABAAAAFgAAAAwAAAAAAAAAVAAAANwAAAAKAAAAcAAAAI8AAAB8AAAAAQAAAADgukGrqrpBCgAAAHAAAAAYAAAATAAAAAQAAAAJAAAAcAAAAJEAAAB9AAAAfAAAAFMAaQBnAG4AZQBkACAAYgB5ADoAIABPAHMAcwBpAGEAbgBFAGQAcwB0AHIA9gBtAAYAAAADAAAABwAAAAcAAAAGAAAABwAAAAMAAAAHAAAABQAAAAMAAAADAAAACQAAAAUAAAAFAAAAAwAAAAYAAAAHAAAABgAAAAcAAAAFAAAABAAAAAQAAAAHAAAACQAAABYAAAAMAAAAAAAAACUAAAAMAAAAAgAAAA4AAAAUAAAAAAAAABAAAAAUAAAA";

            Assert.AreEqual(invalidImage, originalInvalidImage);
        }

        [TestMethod]
        public void TestTextLength()
        {
            var inValidTemplate = new SignatureLineTemplateEmf();
            inValidTemplate.InsertInvalidRecords();

            string testText = "IHaveAVeryVeryVeryVerylon";
            inValidTemplate.SignText = testText;
            Assert.AreEqual(inValidTemplate.signTextObject.Text, testText);

            testText = "IHaveAVeryVeryVeryVerylong";
            inValidTemplate.SignText = testText;
            Assert.AreEqual(inValidTemplate.signTextObject.Text, "IHaveAVeryVeryVeryVerylo...");

            testText = "IHaveAVeryVeryVeryVerylonggggggggggggggggggggggggggggggggggggggggg";
            inValidTemplate.SignText = testText;
            Assert.AreEqual(inValidTemplate.signTextObject.Text, "IHaveAVeryVeryVeryVerylo...");

            testText = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLM";
            inValidTemplate.SuggestedSigner = testText;
            Assert.AreEqual(inValidTemplate.suggestedSignerObject.Text, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLM");

            inValidTemplate.SuggestedTitle = testText;
            Assert.AreEqual(inValidTemplate.suggestedTitleObject.Text, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLM");

            testText += "N";
            inValidTemplate.SuggestedSigner = testText;
            Assert.AreEqual(inValidTemplate.suggestedSignerObject.Text, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKL...");

            inValidTemplate.SuggestedTitle = testText;
            Assert.AreEqual(inValidTemplate.suggestedTitleObject.Text, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKL...");

            testText += "OPQR";
            inValidTemplate.SuggestedSigner = testText;
            Assert.AreEqual(inValidTemplate.suggestedSignerObject.Text, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKL...");

            inValidTemplate.SuggestedTitle = testText;
            Assert.AreEqual(inValidTemplate.suggestedTitleObject.Text, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKL...");
        }

        [TestMethod]
        public void SavingEmptyPartShouldCreateFileAndNotThrow()
        {
            string partURI = @"/_xmlsignatures/origin.sigs";
            var partUri = new Uri(partURI, UriKind.Relative);

            using (ExcelPackage package = OpenPackage("DigSig_EmptyPart.xlsx", true))
            {
                package.Workbook.Worksheets.Add("newWorksheet");
                var part = package.ZipPackage.CreatePart(partUri, ContentTypes.signatureOrigin);
                var stream = part.GetStream();
                stream.Write([], 0, 0);
                part.CreateRelationship("sig1.xml", TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature");

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage("DigSig_EmptyPart.xlsx"))
            {
                var wb = package.Workbook;

                bool partExists = wb._package.ZipPackage.PartExists(partUri);
                Assert.IsFalse(partExists);
            }
        }

        [TestMethod]
        public void SignSave()
        {
            using (var pck = OpenPackage("generatedSignedEmpty.xlsx", true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("emptyWorksheet");

                var cert = GetSelfCert();
                var digSig = wb.DigitialSignatures.AddSignature(cert);

                SaveAndCleanup(pck);
            }
        }

        //Normalize
        //Canonize
        //Transform
        //Hash
        //Read as string

        [TestMethod]
        public void SignSaveTemplateSimple()
        {
            using (var pck = OpenTemplatePackage("simpleDoc.xlsx"))
            {
                var wb = pck.Workbook;
                wb.FullCalcOnLoad = false;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1]);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void SignSaveTemplateEmpty()
        {
            using (var pck = OpenTemplatePackage("UnsignedWBEmpty.xlsx"))
            {
                RSACryptoServiceProvider rsaKey = new();

                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1]);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void SignFileExternal()
        {
            using (var pck = OpenTemplatePackage("LinkExternalSign.xlsx"))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1]);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void AddComment()
        {
            using (var pck = OpenPackage("CommentTest.xlsx", true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("CommentWs");

                ws.Cells["A1"].AddComment("Do Something about this", "ossian");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void AddImage()
        {
            using (var pck = OpenPackage("ImageTest.xlsx", true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("imageWs");

                var pic = ws.Drawings.AddPicture("Landscape", new FileInfo(@"C:\Users\OssianEdström\Pictures\webp.jpg"));
                pic.SetPosition(2, 0, 1, 0);

                SaveAndCleanup(pck);
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

            using (var pck = OpenTemplatePackage("combineddatareport.xlsx"))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitialSignatures.AddSignature(cert);
                var info = digSig.Details;

                info.SignerRoleTitle = title;
                info.Address1 = address;
                info.Address2 = address2;
                info.ZIPorPostalCode = ZIPorPostalCode;
                info.City = city;
                info.CountryOrRegion = CountryOrRegion;
                info.StateOrProvince = StateOrProvince;

                SaveAndCleanup(pck);
            }
            using (var pck = OpenPackage("combineddatareport.xlsx"))
            {
                var wb = pck.Workbook;
                var signerInformation = wb.DigitialSignatures[0].Details;
                Assert.AreEqual(title, signerInformation.SignerRoleTitle);
                Assert.AreEqual(address, signerInformation.Address1);
                Assert.AreEqual(address2, signerInformation.Address2);
                Assert.AreEqual(ZIPorPostalCode, signerInformation.ZIPorPostalCode);
                Assert.AreEqual(city, signerInformation.City);
                Assert.AreEqual(CountryOrRegion, signerInformation.CountryOrRegion);
                Assert.AreEqual(StateOrProvince, signerInformation.StateOrProvince);
            }
        }

        [TestMethod]
        public void SignSaveFileWithLOTSOfData()
        {
            using (var pck = OpenTemplatePackage("s350.xlsm"))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1]);

                SaveAndCleanup(pck);
            }
        }

        //Interestingly enough. Excel gets invalid signature when EXCEL tries to save this.
        //We do too
        //[TestMethod]
        //public void SignSaveFileWithLOTSOfData2()
        //{
        //    using (var pck = OpenTemplatePackage("S610.xlsx"))
        //    {
        //        var wb = pck.Workbook;

        //        wb.FullCalcOnLoad = false;

        //        X509Store store = new X509Store(StoreLocation.CurrentUser);
        //        store.Open(OpenFlags.ReadOnly);
        //        var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1]);

        //        SaveAndCleanup(pck);
        //    }
        //}


        [TestMethod]
        public void SignSaveFileWithData()
        {
            using (var pck = OpenTemplatePackage("StackedLabelsMoveNineThree.xlsx"))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1]);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ReadSignedFile()
        {
            using (ExcelPackage pck = OpenTemplatePackage("simpleDocExcelSigned.xlsx"))
            {
                var wb = pck.Workbook;
                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void CreateDigSigSHA512()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLineSHA512.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "ASigner";

                var digSig = ws.Workbook.DigitialSignatures.AddSignature(GetSelfCert(), CommitmentType.CreatedAndApproved, "TestingSignatureLine");
                var info = digSig.Details;

                digSig.SetDigestMethod(DigitalSignatureHashAlgorithm.SHA512);

                Assert.AreEqual("http://www.w3.org/2001/04/xmldsig-more#rsa-sha512", digSig._signatureMethod);
                Assert.AreEqual("http://www.w3.org/2001/04/xmlenc#sha512", digSig._digestMethod);

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZIPorPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                sLine.SignWithExistingText(digSig, "ASigner");

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage("DigSig_SignatureLineSHA512.xlsx"))
            {
                var wb = package.Workbook;

                var digSig = wb.DigitialSignatures[0];

                //Ensure it is read correctly:
                Assert.AreEqual("http://www.w3.org/2001/04/xmldsig-more#rsa-sha512", digSig._signatureMethod);
                Assert.AreEqual("http://www.w3.org/2001/04/xmlenc#sha512", digSig._digestMethod);

                SaveAndCleanup(package);
            }
        }
    }
}
