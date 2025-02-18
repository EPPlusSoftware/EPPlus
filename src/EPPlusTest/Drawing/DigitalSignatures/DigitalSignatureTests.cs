using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security.Cryptography;
using System.Xml;
using System;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using System.Text;
using OfficeOpenXml.Utils;
using System.Runtime.ConstrainedExecution;
using System.IO;

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

        //Alternate method for if you have a valid cert locally already
        //X509Certificate2 GetSelfCert()
        //{
        //    X509Store store = new X509Store(StoreLocation.CurrentUser);
        //    store.Open(OpenFlags.ReadOnly);
        //    return store.Certificates[1];
        //}

            [ClassInitialize]
        public static void Init(TestContext context)
        {
            CreatePathIfNotExists(_worksheetPath + SubFolder);
        }

        [ClassCleanup]
        public static void Cleanup()
        {
        }

        void FillInfoWithString(AdditionalSignatureInfo info, string s)
        {
            info.SignerRoleTitle = s;
            info.Address1 = s;
            info.Address2 = s;
            info.ZipOrPostalCode = s;
            info.City = s;
            info.CountryOrRegion = s;
            info.StateOrProvince = s;
        }

        [TestMethod]
        public void CreateDigitalSignatureAndReadIt()
        {
            string fileName = $"{SubFolder}InvisibleSignature.xlsx";
            var cert = GetSelfCert();

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                wb.Calculate();

                var test = package.Workbook.FullCalcOnLoad;

                var digSig = ws.Workbook.DigitialSignatures.Add(cert);
                digSig.CommitmentTyping = CommitmentType.CreatedAndApproved;
                digSig.PurposeForSigning = "TestingSignatureLine";

                var info = digSig.Details;

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZipOrPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                SaveAndCleanup(package);
                Assert.IsTrue(digSig.IsValid);
            }
            using (ExcelPackage package = OpenPackage(fileName))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets[0];
                var digSig = wb.DigitialSignatures[0];
                digSig.Certificate = cert;

                var info = digSig.Details;
                Assert.AreEqual("A Title", info.SignerRoleTitle);
                Assert.AreEqual("Some", info.Address1);
                Assert.AreEqual("Where", info.Address2);
                Assert.AreEqual("Over", info.ZipOrPostalCode);
                Assert.AreEqual("The", info.City);
                Assert.AreEqual("Rainbow", info.CountryOrRegion);
                Assert.AreEqual("WayUpHigh", info.StateOrProvince);

                SaveAndCleanup(package);
                Assert.IsTrue(digSig.IsValid);
            }
        }

        [TestMethod]
        public void ReadCommitmentTypeAndTypeQualifier()
        {
            string fileName = "DigSig_FullSignatureAndLine.xlsx";

            using (ExcelPackage package = OpenTemplatePackage(fileName))
            {
                var wb = package.Workbook;

                var digSig = wb.DigitialSignatures[0];
                digSig.Certificate = GetSelfCert();
                Assert.AreEqual(CommitmentType.Approved, digSig.CommitmentTyping);
                Assert.AreEqual("MyPurposeIsMyOwn", digSig.PurposeForSigning);

                package.SaveAs(GetOutputFile(SubFolder, fileName));
            }
        }
        [TestMethod]
        public void ReadCommitmentTypeAndTypeQualifierWhenNone()
        {
            string fileName = "DigSig_FullSignatureAndLineNone.xlsx";

            using (ExcelPackage package = OpenTemplatePackage(fileName))
            {
                var wb = package.Workbook;

                var digSig = wb.DigitialSignatures[0];
                digSig.Certificate = GetSelfCert();

                Assert.AreEqual(CommitmentType.None, digSig.CommitmentTyping);
                Assert.AreEqual("MyPurposeIsMyOwn", digSig.PurposeForSigning);

                package.SaveAs(GetOutputFile(SubFolder, fileName));
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
        public void SignSignedWorkbook()
        {
            string fileName = $"{SubFolder}DoubleSigned.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("newWs");

                wb.DigitialSignatures.Add(GetSelfCert());

                SaveAndCleanup(package);
            }
            //Sign an already signed workbook should add and result in two valid signatures.
            using (ExcelPackage package = OpenPackage(fileName))
            {
                var wb = package.Workbook;

                var digSig = wb.DigitialSignatures.Add(GetSelfCert());
                digSig.CommitmentTyping = CommitmentType.CreatedAndApproved;
                digSig.PurposeForSigning = "DoubleSigning";

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage(fileName))
            {
                var wb = package.Workbook;

                //Ensure both are valid and tests enumerator.
                int i = 0;
                foreach (var sig in wb.DigitialSignatures)
                {
                    Assert.IsTrue(sig.IsValid);
                    i += 1;
                }
                Assert.AreEqual(2, i);
            }
        }


        [TestMethod]
        public void CounterSignAfterChanges()
        {
            string fileName = $"{SubFolder}CounterSigned.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("newWs");

                var digSig = wb.DigitialSignatures.Add(GetSelfCert());
                digSig.CommitmentTyping = CommitmentType.CreatedAndApproved;
                digSig.PurposeForSigning = "Counter-signing";

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage(fileName))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets[0];

                ws.Cells["A1"].Value = 5;

                var sig = wb.DigitialSignatures[0];
                sig.Certificate = GetSelfCert();

                SaveAndCleanup(package);
            }
        }


        [TestMethod]
        public void CreateDigitalSignatureLineAndSignIt()
        {
            string fileName = $"{SubFolder}DigSig_SignatureLine.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.SignatureLines.Add();
                sLine.Signer = "ASigner";

                var digSig = wb.DigitialSignatures.Add(GetSelfCert());
                digSig.CommitmentTyping = CommitmentType.CreatedAndApproved;
                digSig.PurposeForSigning = "TestingSignatureLine";

                var info = digSig.Details;

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZipOrPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                digSig.SignatureLine = sLine;

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void SavingEmptyPartShouldCreateFileAndNotThrow()
        {
            string partURI = @"/_xmlsignatures/origin.sigs";
            var partUri = new Uri(partURI, UriKind.Relative);

            string fileName = $"{SubFolder}DigSig_EmptyPart.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                package.Workbook.Worksheets.Add("newWorksheet");
                var part = package.ZipPackage.CreatePart(partUri, ContentTypes.signatureOrigin);
                var stream = part.GetStream();
                stream.Write([], 0, 0);
                part.CreateRelationship("sig1.xml", TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature");

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage(fileName))
            {
                var wb = package.Workbook;

                bool partExists = wb._package.ZipPackage.PartExists(partUri);
                Assert.IsFalse(partExists);
            }
        }

        [TestMethod]
        public void SignSave()
        {
            string fileName = $"{SubFolder}generatedSignedEmpty.xlsx";
            var cert = GetSelfCert();

            using (var pck = OpenPackage(fileName, true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("emptyWorksheet");

                wb.Calculate();

                var digSig = wb.DigitialSignatures.Add(cert);

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
            string fileName = $"simpleDoc.xlsx";

            using (var pck = OpenTemplatePackage(fileName))
            {
                var wb = pck.Workbook;
                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitialSignatures.Add(cert);

                pck.SaveAs(GetOutputFile(SubFolder, fileName));
            }
        }

        [TestMethod]
        public void SignSaveTemplateEmpty()
        {
            string fileName = $"UnsignedWBEmpty.xlsx";

            using (var pck = OpenTemplatePackage("UnsignedWBEmpty.xlsx"))
            {
                RSACryptoServiceProvider rsaKey = new();

                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitialSignatures.Add(cert);

                pck.SaveAs(GetOutputFile(SubFolder, fileName));
            }
        }

        [TestMethod]
        public void SignFileExternal()
        {
            string fileName = $"LinkExternalSign.xlsx";

            using (var pck = OpenTemplatePackage(fileName))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitialSignatures.Add(cert);

                pck.SaveAs(GetOutputFile(SubFolder, fileName));
            }
        }

        [TestMethod]
        public void AddComment()
        {
            string fileName = $"{SubFolder}CommentTest.xlsx";
            var cert = GetSelfCert();

            //var key = cert.GetRSAPrivateKey();

            using (var pck = OpenPackage(fileName, true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("CommentWs");

                //wb.Calculate();

                //ws.Cells["A1"].AddComment("Do Something about this", "ossian");
                var sigLine = ws.SignatureLines.Add();
                sigLine.SignWithText(cert, "ASigner");

                var test = wb.DigitialSignatures[0].Certificate.GetRSAPrivateKey();
                var test2 = wb.DigitialSignatures[0].Certificate.GetRSAPrivateKey();

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

            string fileName = $"combineddatareport.xlsx";

            using (var pck = OpenTemplatePackage(fileName))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitialSignatures.Add(cert);
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
                var signerInformation = wb.DigitialSignatures[0].Details;
                Assert.AreEqual(title, signerInformation.SignerRoleTitle);
                Assert.AreEqual(address, signerInformation.Address1);
                Assert.AreEqual(address2, signerInformation.Address2);
                Assert.AreEqual(ZIPorPostalCode, signerInformation.ZipOrPostalCode);
                Assert.AreEqual(city, signerInformation.City);
                Assert.AreEqual(CountryOrRegion, signerInformation.CountryOrRegion);
                Assert.AreEqual(StateOrProvince, signerInformation.StateOrProvince);
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
        public void SignSaveFileWithData()
        {
            string fileName = "StackedLabelsMoveNineThree.xlsx";

            using (var pck = OpenTemplatePackage(fileName))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                var cert = GetSelfCert();
                var digSig = wb.DigitialSignatures.Add(cert);

                pck.SaveAs(GetOutputFile(SubFolder, fileName));
            }
        }

        [TestMethod]
        public void ReadSignedFile()
        {
            string fileName = "simpleDocExcelSigned.xlsx";
            using (ExcelPackage pck = OpenTemplatePackage(fileName))
            {
                var wb = pck.Workbook;
                pck.SaveAs(GetOutputFile(SubFolder, fileName));
            }
        }

        public void TestHashAlgorithm(string signatureMethod, string digestMethod, DigitalSignatureHashAlgorithm algorithm)
        {
            var name = Enum.GetName(typeof(DigitalSignatureHashAlgorithm), algorithm);

            string fileName = $"{SubFolder}DigSig_SignatureLine{name}.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.SignatureLines.Add();
                sLine.Signer = "ASigner";

                var digSig = wb.DigitialSignatures.Add(GetSelfCert());
                digSig.CommitmentTyping = CommitmentType.CreatedAndApproved;
                digSig.PurposeForSigning = "TestingSignatureLine";

                var info = digSig.Details;

                digSig.SetDigestMethod(algorithm);

                Assert.AreEqual(signatureMethod, digSig._signatureMethod);
                Assert.AreEqual(digestMethod, digSig._digestMethod);

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZipOrPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                var signature = sLine.SignWithText(GetSelfCert(), "ASigner");
                signature.Details.Address1 = "Address";
                signature.CommitmentTyping = CommitmentType.Approved;

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage(fileName))
            {
                var wb = package.Workbook;

                var digSig = wb.DigitialSignatures[0];

                //Ensure it is read correctly:
                Assert.AreEqual(signatureMethod, digSig._signatureMethod);
                Assert.AreEqual(digestMethod, digSig._digestMethod);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void CreateDigSigSHA1()
        {
            TestHashAlgorithm("http://www.w3.org/2000/09/xmldsig#rsa-sha1", "http://www.w3.org/2000/09/xmldsig#sha1", DigitalSignatureHashAlgorithm.SHA1);
        }
        [TestMethod]
        public void CreateDigSigSHA256()
        {
            TestHashAlgorithm("http://www.w3.org/2001/04/xmldsig-more#rsa-sha256", "http://www.w3.org/2001/04/xmlenc#sha256", DigitalSignatureHashAlgorithm.SHA256);
        }
        [TestMethod]
        public void CreateDigSigSHA384()
        {
            TestHashAlgorithm("http://www.w3.org/2001/04/xmldsig-more#rsa-sha384", "http://www.w3.org/2001/04/xmldsig-more#sha384", DigitalSignatureHashAlgorithm.SHA384);
        }

        [TestMethod]
        public void CreateDigSigSHA512()
        {
            TestHashAlgorithm("http://www.w3.org/2001/04/xmldsig-more#rsa-sha512", "http://www.w3.org/2001/04/xmlenc#sha512", DigitalSignatureHashAlgorithm.SHA512);
        }


        [TestMethod]
        public void TwoDifferentHashAlgsOnDifferentSignaturesShouldOnlyBeTheLastSet()
        {
            string fileName = $"{SubFolder}DoubleHashes.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureWs");

                var signature = wb.DigitialSignatures.Add(GetSelfCert());
                signature.SetDigestMethod(DigitalSignatureHashAlgorithm.SHA384);

                var signature2 = wb.DigitialSignatures.Add(GetSelfCert());
                signature2.SetDigestMethod(DigitalSignatureHashAlgorithm.SHA512);

                SaveAndCleanup(package);
            }
        }


        [TestMethod]
        public void CreatingSignatureLineFilledAfterHash()
        {
            string fileName = $"{SubFolder}DoubleHashesNoSet.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureWs");

                var signature = wb.DigitialSignatures.Add(GetSelfCert());
                signature.SetDigestMethod(DigitalSignatureHashAlgorithm.SHA384);

                var sLine = ws.SignatureLines.Add();

                sLine.AllowComments = true;
                sLine.SigningInstructions = "Some instructions";
                sLine.Signer = "Eric";
                sLine.Title = "The Eternal";
                sLine.Email = "TheEternal@SufferingPunishment.se";
                sLine.AllowComments = true;
                sLine.ShowSignDate = true;

                sLine.SignWithText(GetSelfCert(), "SomeText");

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void SetFaultyFromToRow()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("AWs");

                var sLine = ws.SignatureLines.Add();
                sLine.From.Row = 5;
                sLine.To.Row = 4;

                package.Save();
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void SetFaultyFromToColumn()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("AWs");

                var sLine = ws.SignatureLines.Add();
                sLine.From.Column = 5;
                sLine.To.Column = 4;

                package.Save();
            }
        }

        [TestMethod]
        public void EnsureTextIsEscaped()
        {
            string fileName = $"{SubFolder}StrangeSymbolsEscapedSignedSignatureLine.xlsx";
            var symbols = "\"&%/Stuff���}=``#�\"<>\"An Example\"";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.SignatureLines.Add();
                sLine.Signer = symbols;
                sLine.Title = symbols;
                sLine.Email = symbols;
                sLine.SigningInstructions = symbols;
                sLine.AllowComments = true;

                var digSig = wb.DigitialSignatures.Add(GetSelfCert());
                digSig.CommitmentTyping = CommitmentType.CreatedAndApproved;
                digSig.PurposeForSigning = symbols;

                var info = digSig.Details;
                FillInfoWithString(info, symbols);

                var signature = sLine.SignWithText(GetSelfCert(), symbols);
                signature.PurposeForSigning = symbols;
                signature.CommitmentTyping = CommitmentType.Approved;

                FillInfoWithString(signature.Details, symbols);

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage(fileName))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets[0];

                var sigLine = ws.SignatureLines[0].AsSignatureLine;

                Assert.AreEqual(symbols, sigLine.Signer);
                Assert.AreEqual(symbols, sigLine.Title);
                Assert.AreEqual(symbols, sigLine.Email);
                Assert.AreEqual(symbols, sigLine.SigningInstructions);

                var digSig = sigLine.DigitalSignature;

                Assert.AreEqual(symbols, digSig.Details.SignerRoleTitle);
                Assert.AreEqual(symbols, digSig.Details.Address1);
                Assert.AreEqual(symbols, digSig.Details.Address2);
                Assert.AreEqual(symbols, digSig.Details.ZipOrPostalCode);
                Assert.AreEqual(symbols, digSig.Details.City);
                Assert.AreEqual(symbols, digSig.Details.CountryOrRegion);
                Assert.AreEqual(symbols, digSig.Details.StateOrProvince);
            }
        }

        [TestMethod]
        public void RemoveDigitalSignatures()
        {
            string fileName = $"{SubFolder}AddedAndRemovedSignatures.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("emptySignatures");

                var digSig = wb.DigitialSignatures.Add(GetSelfCert());
                digSig.Details.Address1 = "SomeAddress";

                ws.Cells["A1"].Value = "53";

                wb.DigitialSignatures.Remove(digSig);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void RemoveDigitalSignaturesAddAgain()
        {
            string fileName = $"{SubFolder}AddedAndRemovedSignaturesAddedAgain.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("emptySignatures");

                var digSig = wb.DigitialSignatures.Add(GetSelfCert());
                digSig.Details.Address1 = "SomeAddress";

                ws.Cells["A1"].Value = "53";

                wb.DigitialSignatures.Remove(digSig);

                SaveAndCleanup(package);
            }

            using (ExcelPackage pck = OpenPackage(fileName))
            {

                var wb = pck.Workbook;
                var ws = pck.Workbook.Worksheets[0];
                var digSig = wb.DigitialSignatures.Add(GetSelfCert());

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void RemoveDigitalSignaturesAddAgainNoResave()
        {
            string fileName = $"{SubFolder}AddedAndRemovedSignaturesAddedAgainNoResave.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("emptySignatures");

                var digSig = wb.DigitialSignatures.Add(GetSelfCert());
                digSig.Details.Address1 = "SomeAddress";

                ws.Cells["A1"].Value = "53";

                wb.DigitialSignatures.Remove(digSig);

                var digSigNew = wb.DigitialSignatures.Add(GetSelfCert());
                digSigNew.Details.Address1 = "Another address";

                var digSigRead = wb.DigitialSignatures[0];

                Assert.AreEqual("Another address", digSigRead.Details.Address1);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CheckingDigitalSignaturesWhenNotExists()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("emptySignatures");
                var digSigCollection = wb.DigitialSignatures;
            }
        }
    }
}
