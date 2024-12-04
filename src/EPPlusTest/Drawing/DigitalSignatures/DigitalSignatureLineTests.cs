using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Drawing;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml.DigitalSignatures;

namespace EPPlusTest.Drawing.DigitalSignatures
{
    [TestClass]
    public class DigitalSignatureLineTests : TestBase
    {
        const string SubFolder = "DigitalSignatureLines\\";

        X509Certificate2 GetSelfCert()
        {
            var requestedCert = new CertificateRequest("cn=SelfSignCert", RSA.Create(), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var finalCert = requestedCert.CreateSelfSigned(DateTimeOffset.Now, DateTimeOffset.Now.AddMinutes(5));

            var certPrivate = finalCert.Export(X509ContentType.Pfx);
            var certPublic = finalCert.Export(X509ContentType.Cert);
            return new X509Certificate2(certPrivate, "", X509KeyStorageFlags.Exportable);
        }

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            CreatePathIfNotExists(_worksheetPath + SubFolder);
        }

        [ClassCleanup]
        public static void Cleanup()
        {

        }

        [TestMethod]
        public void CreateAndReadDefaultSignatureLine()
        {
            var wsName = "SignatureLineWorksheet";

            using (ExcelPackage package = OpenPackage($"{SubFolder}UnsignedSignatureLine.xlsx"))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add(wsName);

                var sline = ws.AddSignatureLine();

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage($"{SubFolder}UnsignedSignatureLine.xlsx"))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.GetByName(wsName);

                var sigLine = ws.SignatureLines[0];

                Assert.AreEqual(sigLine.SignatureLineType, eSignatureLineType.SignatureLine);
                Assert.AreEqual("", sigLine.Signer);
                Assert.AreEqual("", sigLine.Title);
                Assert.AreEqual("Before signing this document, verify that the content you are signing is correct.", sigLine.SigningInstructions);
                Assert.AreEqual(true, sigLine.ShowSignDate);
                Assert.AreEqual(false, sigLine.AllowComments);
                Assert.AreEqual("Microsoft Office Signature Line...", sigLine.AlternativeText);
                Assert.AreEqual("{00000000-0000-0000-0000-000000000000}", sigLine.ProvID);
            }
        }

        [TestMethod]
        public void CreateAndReadDefaultSignatureLineStamp()
        {
            var wsName = "SignatureLineStamps";

            using (ExcelPackage package = OpenPackage($"{SubFolder}UnsignedSignatureLine.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add(wsName);

                var slineStamp = ws.AddSignatureLineStamp();
                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage($"{SubFolder}UnsignedSignatureLine.xlsx"))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.GetByName(wsName);

                var sigLine = ws.SignatureLines[0];

                Assert.AreEqual(sigLine.SignatureLineType, eSignatureLineType.Stamp);
                Assert.AreEqual("", sigLine.Signer);
                Assert.AreEqual("", sigLine.Title);
                Assert.AreEqual("Before signing this document, verify that the content you are signing is correct.", sigLine.SigningInstructions);
                Assert.AreEqual(true, sigLine.ShowSignDate);
                Assert.AreEqual(false, sigLine.AllowComments);
                Assert.AreEqual("Stamp Signature Line...", sigLine.AlternativeText);
                Assert.AreEqual("{000CD6A4-0000-0000-C000-000000000046}", sigLine.ProvID);
            }
        }

        [TestMethod]
        public void CreateAndReadMultipleSignatureLines()
        {
            var wsName = "MultipleEmptySignatureLines";
            string Signer = "Someone";
            string Title = "WithATitle";
            string SigningInstructions = "NewInstructions";
            bool ShowSignDate = false;
            bool AllowComments = true;
            string AlternativeText = "Alt text";
            var signatureImage = new ExcelImage(GetResourceFile("Code.bmp"));

            using (ExcelPackage package = OpenPackage($"{SubFolder}UnsignedSignatureLine.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add(wsName);

                var sline = ws.AddSignatureLineStamp();

                var sLine2 = ws.AddSignatureLineStamp();
                var sLine3 = ws.AddSignatureLine();
                var sLine4 = ws.AddSignatureLineStamp();
                var sLine5 = ws.AddSignatureLine();

                sLine2.From.Column = 3;
                sLine2.To.Column = 5;

                sLine3.From.Row = 9;
                sLine3.To.Row = 9 + 6;
                sLine3.From.Column = 5;
                sLine3.To.Column = 5 + 4;

                sLine3.Signer = Signer;
                sLine3.Title = Title;
                sLine3.SigningInstructions = SigningInstructions;
                sLine3.ShowSignDate = ShowSignDate;
                sLine3.AllowComments = AllowComments;
                sLine3.AlternativeText = AlternativeText;

                sLine4.From.Column = 10;
                sLine4.To.Column = 12;

                sLine4.Signer = sLine3.Signer;
                sLine4.Title = sLine3.Title;
                sLine4.SigningInstructions = sLine3.SigningInstructions;
                sLine4.ShowSignDate = sLine3.ShowSignDate;
                sLine4.AllowComments = sLine3.AllowComments;
                sLine4.AlternativeText = sLine3.AlternativeText;

                sLine5.From.Row = 9;
                sLine5.To.Row = 9 + 6;
                sLine5.From.Column = 10;
                sLine5.To.Column = 14;

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage($"{SubFolder}UnsignedSignatureLine.xlsx"))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.GetByName(wsName);

                Assert.AreEqual(ws.SignatureLines[0].SignatureLineType, eSignatureLineType.Stamp);
                Assert.AreEqual(ws.SignatureLines[1].SignatureLineType, eSignatureLineType.Stamp);

                var sigLine = ws.SignatureLines[2];

                Assert.AreEqual(sigLine.SignatureLineType, eSignatureLineType.SignatureLine);
                Assert.AreEqual(Signer, sigLine.Signer);
                Assert.AreEqual(Title, sigLine.Title);
                Assert.AreEqual(SigningInstructions, sigLine.SigningInstructions);
                Assert.AreEqual(ShowSignDate, sigLine.ShowSignDate);
                Assert.AreEqual(AllowComments, sigLine.AllowComments);
                Assert.AreEqual(AlternativeText, sigLine.AlternativeText);
                Assert.AreEqual("{00000000-0000-0000-0000-000000000000}", sigLine.ProvID);

                Assert.AreEqual(9, sigLine.From.Row);
                Assert.AreEqual(9 + 6, sigLine.To.Row);
                Assert.AreEqual(5, sigLine.From.Column);
                Assert.AreEqual(9, sigLine.To.Column);

                Assert.AreEqual(9, ws.SignatureLines[4].From.Row);
                Assert.AreEqual(9 + 6, ws.SignatureLines[4].To.Row);
                Assert.AreEqual(10, ws.SignatureLines[4].From.Column);
                Assert.AreEqual(14, ws.SignatureLines[4].To.Column);

                var sigStamp = ws.SignatureLines[3];
                Assert.AreEqual(sigStamp.SignatureLineType, eSignatureLineType.Stamp);
                Assert.AreEqual(Signer, sigStamp.Signer);
                Assert.AreEqual(Title, sigStamp.Title);
                Assert.AreEqual(SigningInstructions, sigStamp.SigningInstructions);
                Assert.AreEqual(ShowSignDate, sigStamp.ShowSignDate);
                Assert.AreEqual(AllowComments, sigStamp.AllowComments);
                Assert.AreEqual(AlternativeText, sigStamp.AlternativeText);
                Assert.AreEqual("{000CD6A4-0000-0000-C000-000000000046}", sigStamp.ProvID);

                Assert.AreEqual(10, sigStamp.From.Column);
                Assert.AreEqual(12, sigStamp.To.Column);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CopyShouldIgnoreSignatureLines()
        {
            var wsName = "originalSheet";
            using (ExcelPackage package = OpenPackage($"{SubFolder}UnsignedSignatureLine.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add(wsName);
                ws.AddSignatureLine();

                var copiedWs = wb.Worksheets.Copy(wsName, "CopiedSheet");

                Assert.AreEqual(0, copiedWs.SignatureLines.Count());
                Assert.AreEqual(0, copiedWs._vmlDrawings.Count);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void SignSignatureLine()
        {
            var signatureImage = new ExcelImage(GetResourceFile("Code.bmp"));

            using (ExcelPackage package = OpenPackage($"{SubFolder}SignedSignatureLines.xlsx", true))
            {
                var wb = package.Workbook;
                var sSline = package.Workbook.Worksheets.Add("SignedSignatureLine");
                var manysSlines = package.Workbook.Worksheets.Add("ManySignedSignatureLine");
                var sSlineStamp = package.Workbook.Worksheets.Add("SignedSignatureLineStamp");
                var manySlineStamps = package.Workbook.Worksheets.Add("ManySignedSignatureLineStamp");

                var cert = GetSelfCert();

                var sLine = sSline.AddSignatureLine();
                sLine.Sign(cert, signatureImage);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CreateSignAndResaveSigLine()
        {
            var signatureImage = new ExcelImage(GetResourceFile("Code.bmp"));

            using (ExcelPackage package = OpenPackage($"{SubFolder}SignedSignatureLinesResaved.xlsx", true))
            {
                var wb = package.Workbook;
                var sSline = package.Workbook.Worksheets.Add("SignedSignatureLine");

                var sLine = sSline.AddSignatureLine();

                var cert = GetSelfCert();
                sLine.Sign(cert, signatureImage);

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage($"{SubFolder}SignedSignatureLinesResaved.xlsx"))
            {
                var wb = package.Workbook;

                var sig = wb.DigitialSignatures[0];

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void ReadSline()
        {
            string origImage;
            string origValid;
            string origInvalid;
            string sharedName = "Excel";

            byte[] originalImgBytes;

            using (var pck = OpenTemplatePackage("BmpImage.xlsx"))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets[0];

                var sLine = ws.SignatureLines[0];
                var digSig = wb.DigitialSignatures[0];

                originalImgBytes = sLine.SignatureImage.ImageBytes;

                origImage = digSig.SigLnImage;
                origValid = digSig.ValidSigLnImage;
                origInvalid = digSig.InvalidSigLnImg;

                var imageOnly = GetOutputFile(SubFolder, $"{sharedName}Image.emf");
                var noSigOnlyTemplate = GetOutputFile(SubFolder, $"{sharedName}image1NoSig.emf");
                var validImg = GetOutputFile(SubFolder, $"{sharedName}Valid.emf");
                var invalidImg = GetOutputFile(SubFolder, $"{sharedName}Invalid.emf");

                DecodeAndSaveEmf(sLine.SigLnImage, imageOnly.FullName);
                var emptyTemplatePart = pck.Workbook._package.ZipPackage.GetPart(new Uri("xl\\media\\image1.emf", UriKind.Relative));
                var imageStream = (MemoryStream)emptyTemplatePart.GetStream(FileMode.Open, FileAccess.Read);
                var bytes = imageStream.ToArray();
                File.WriteAllBytes(noSigOnlyTemplate.FullName, bytes);

                DecodeAndSaveEmf(sLine.ValidSigLnImage, validImg.FullName);
                DecodeAndSaveEmf(sLine.InvalidSigLnImg, invalidImg.FullName);

                DecodeAndSaveEmf(origImage, GetOutputFile(SubFolder, "SignatureImagePng.emf").FullName);
                DecodeAndSaveEmf(origValid, GetOutputFile(SubFolder, "SignatureImagePngValid.emf").FullName);
                DecodeAndSaveEmf(origInvalid, GetOutputFile(SubFolder, "SignatureImagePngInvalid.emf").FullName);

                SaveAndCleanup(pck);
            }

            string openedOnceImage;
            string openedOnceValid;
            string openedOnceInvalid;
            string digitalSignatureOuterXmlOnceOpened;

            using (var pck = OpenPackage("BmpImage.xlsx"))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets[0];

                var aSline = ws.SignatureLines[0];
                var digSig = wb.DigitialSignatures[0];

                openedOnceImage = digSig.SigLnImage;
                openedOnceValid = digSig.ValidSigLnImage;
                openedOnceInvalid = digSig.InvalidSigLnImg;

                //Epplus should have generated a new signature using our templates
                //Since opening the file in epplus changes some files. (Notably shared strings)
                //The images should be very similar but different.
                Assert.AreNotEqual(origImage, openedOnceImage);
                Assert.AreNotEqual(origValid, openedOnceValid);
                Assert.AreNotEqual(origInvalid, openedOnceInvalid);

                digitalSignatureOuterXmlOnceOpened = digSig.GetOuterXml();

                //The actual .bmp file should be the same.
                Assert.IsTrue(originalImgBytes.SequenceEqual(aSline.SignatureImage.ImageBytes));

                SaveAndCleanup(pck);
            }

            using (var pck = OpenPackage("BmpImage.xlsx"))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets[0];
                var digSig = wb.DigitialSignatures[0];
                var aSline = ws.SignatureLines[0];

                //Opening again with Epplus without changing anything should be the same
                Assert.AreEqual(openedOnceImage, digSig.SigLnImage);
                Assert.AreEqual(openedOnceValid, digSig.ValidSigLnImage);
                Assert.AreEqual(openedOnceInvalid, digSig.InvalidSigLnImg);
                Assert.IsTrue(originalImgBytes.SequenceEqual(aSline.SignatureImage.ImageBytes));
                Assert.AreEqual(digitalSignatureOuterXmlOnceOpened, digSig.GetOuterXml());

                //Changing a value should cause a re-signing on save
                ws.Cells["A1"].Value = "changedValue";

                SaveAndCleanup(pck);
            }

            using (var pck = OpenPackage("BmpImage.xlsx"))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets[0];
                var digSig = wb.DigitialSignatures[0];
                var aSline = ws.SignatureLines[0];

                //Opening again with Epplus should be the same
                Assert.AreEqual(openedOnceImage, digSig.SigLnImage);
                Assert.AreEqual(openedOnceValid, digSig.ValidSigLnImage);
                Assert.AreEqual(openedOnceInvalid, digSig.InvalidSigLnImg);
                Assert.IsTrue(originalImgBytes.SequenceEqual(aSline.SignatureImage.ImageBytes));

                //But the signature itself different as the Sheet1 hash has changed
                Assert.AreNotEqual(digitalSignatureOuterXmlOnceOpened, digSig.GetOuterXml());

                SaveAndCleanup(pck);

                Assert.IsTrue(digSig.IsValid);
            }
        }

        [TestMethod]
        public void VerifySignatureLineEmfs()
        {
            var sharedName = "EpplusSigEmf";
            var fileName = $"{SubFolder}{sharedName}.xlsx";
            var signatureImage = new ExcelImage(GetResourceFile("Code.bmp"));

            using (var pck = OpenPackage(fileName, true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("sLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.SigningInstructions = "These are Instructions";
                sLine.Signer = "ASigner";
                sLine.Title = "SomeDeveloper";
                sLine.Email = "Some@developer.se";
                sLine.AllowComments = true;
                sLine.ShowSignDate = true;

                var cert = GetSelfCert();
                sLine.Sign(cert, signatureImage);

                SaveAndCleanup(pck);
            }

            using (var pck = OpenPackage(fileName))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets[0];
                var signature = wb.DigitialSignatures[0];
                var sLine = ws.SignatureLines[0];

                var imageOnly = GetOutputFile(SubFolder, $"{sharedName}Image.emf");
                var noSigOnlyTemplate = GetOutputFile(SubFolder, $"{sharedName}image1NoSig.emf");
                var validImg = GetOutputFile(SubFolder, $"{sharedName}Valid.emf");
                var invalidImg = GetOutputFile(SubFolder, $"{sharedName}Invalid.emf");

                DecodeAndSaveEmf(sLine.SigLnImage, imageOnly.FullName);
                var emptyTemplatePart = pck.Workbook._package.ZipPackage.GetPart(new Uri("xl\\media\\image1.emf",UriKind.Relative));
                var imageStream = (MemoryStream)emptyTemplatePart.GetStream(FileMode.Open, FileAccess.Read);
                var bytes = imageStream.ToArray();
                File.WriteAllBytes(noSigOnlyTemplate.FullName, bytes);

                DecodeAndSaveEmf(sLine.ValidSigLnImage, validImg.FullName);
                DecodeAndSaveEmf(sLine.InvalidSigLnImg, invalidImg.FullName);

                var imageOnlyEmf = new EmfImage();
                imageOnlyEmf.Read(File.ReadAllBytes(imageOnly.FullName));
                var headerImageOnly = (EMR_HEADER)imageOnlyEmf.records[0];

                //Width And Height of image is over maximum size for template. Should have been adjusted
                Assert.AreNotEqual(signatureImage.Bounds.Width, headerImageOnly.Bounds.Right);
                Assert.AreNotEqual(signatureImage.Bounds.Height, headerImageOnly.Bounds.Bottom);
                Assert.AreEqual(35, headerImageOnly.Bounds.Bottom);
                Assert.AreEqual(205, headerImageOnly.Bounds.Right);


            }
        }


        [TestMethod]
        public void CreateSignAndResaveSigLineFullInfo()
        {
            var signatureImage = new ExcelImage(GetResourceFile("Code.bmp"));
            var fileName = $"{SubFolder}SignedSignatureLinesResavedFullInfo.xlsx";

            using (ExcelPackage package = OpenPackage(fileName, true))
            {
                var wb = package.Workbook;
                var sSline = package.Workbook.Worksheets.Add("SignedSignatureLine");

                var cert = GetSelfCert();

                var sLine = sSline.AddSignatureLine();
                sLine.SigningInstructions = "These are Instructions";
                sLine.Signer = "ASigner";
                sLine.Title = "SomeDeveloper";
                sLine.Email = "Some@developer.se";
                sLine.AllowComments = true;
                sLine.ShowSignDate = true;

                sLine.Sign(cert, signatureImage);

                var digSig = sLine.DigitalSignature;

                var info = digSig.Details;

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZIPorPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                digSig.PurposeForSigning = "My Purpose is My Own";
                digSig.CommitmentTyping = CommitmentType.CreatedAndApproved;

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage(fileName))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets[0];

                var sigLine = ws.SignatureLines[0];

                //Getting signature from signatureline
                var sig = sigLine.DigitalSignature;
                var info = sig.Details;

                //Check signatureline is equal
                Assert.AreEqual("A Title", info.SignerRoleTitle);
                Assert.AreEqual("Some", info.Address1);
                Assert.AreEqual("Where", info.Address2);
                Assert.AreEqual("Over", info.ZIPorPostalCode);
                Assert.AreEqual("The", info.City);
                Assert.AreEqual("Rainbow", info.CountryOrRegion);
                Assert.AreEqual("WayUpHigh", info.StateOrProvince);

                Assert.AreEqual("My Purpose is My Own", sig.PurposeForSigning);
                Assert.AreEqual(CommitmentType.CreatedAndApproved, sig.CommitmentTyping);

                //Check data on the signatureline itself.
                Assert.AreEqual("These are Instructions", sigLine.SigningInstructions);
                Assert.AreEqual("SomeDeveloper", sigLine.Title);
                Assert.AreEqual("Some@developer.se", sigLine.Email);
                Assert.IsTrue(sigLine.AllowComments);
                Assert.IsTrue(sigLine.ShowSignDate);
                Assert.AreEqual(sigLine.SignatureImage.Type, signatureImage.Type);
                Assert.IsTrue(Enumerable.SequenceEqual(sigLine.SignatureImage.ImageBytes, signatureImage.ImageBytes));

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ExcelGeneratedExtractValidInvalid()
        {
            using (ExcelPackage package = OpenTemplatePackage("emfTextTest.xlsx"))
            {
                var wb = package.Workbook;
                var signature = wb.DigitialSignatures[0];

                var validOutput = GetOutputFile("", "ValidTestSignatureEmf.emf");
                var invalidOutput = GetOutputFile("", "InvalidTestSignatureEmf.emf");
                var changedOutput = GetOutputFile("", "ChangedTestStampSignature.emf");

                DecodeAndSaveEmf(signature.ValidSigLnImage, validOutput.FullName);
                DecodeAndSaveEmf(signature.InvalidSigLnImg, invalidOutput.FullName);

                var emfImage = new EmfImage();
                emfImage.Read(validOutput.FullName);
                var records = emfImage.records;

                var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Cast<EMR_EXTTEXTOUTW>().ToList();

                textRecords[1].Text = "Developer";

                emfImage.Save(changedOutput.FullName);
            }
        }

        private void DecodeAndSaveEmf(string base64String, string savePath)
        {
            var decodedBytes = Convert.FromBase64String(base64String);
            File.WriteAllBytes(savePath, decodedBytes);
        }

        [TestMethod]
        public void CreateSignatureLineWithALLWithoutSignatureAndSpecialSymbols()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_ALL.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "Ossian Edström åäö";
                sLine.Title = "#Maker \"Quotation`¨'m!";
                sLine.Email = "Example@Site.com";
                sLine.SigningInstructions = "Hey please sign this because x and y so it will be z";
                sLine.AllowComments = true;
                sLine.ShowSignDate = true;

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_ALL.xlsx"))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets[0];

                var sLine = ws.SignatureLines[0];
                Assert.AreEqual("Ossian Edström åäö", sLine.Signer);
                Assert.AreEqual("#Maker \"Quotation`¨'m!", sLine.Title);
                Assert.AreEqual("Example@Site.com", sLine.Email);
                Assert.AreEqual("Hey please sign this because x and y so it will be z", sLine.SigningInstructions);
                Assert.IsTrue(sLine.AllowComments);
                Assert.IsTrue(sLine.ShowSignDate);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ReadDigitalSignatureLineStamp()
        {
            using (ExcelPackage package = OpenTemplatePackage("FullStamp.xlsx"))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets[0];

                var signatures = package.Workbook.DigitialSignatures;

                DecodeAndSaveEmf(signatures[0].ValidSigLnImage, GetOutputFile("", "validStamp.emf").FullName);
                DecodeAndSaveEmf(signatures[0].InvalidSigLnImg, GetOutputFile("", "invalidStamp.emf").FullName);
            }
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void CreateSignatureLineChangeImage()
        {
            var signatureImage = new ExcelImage(GetResourceFile("Code.bmp"));

            using (ExcelPackage package = OpenPackage("CreateSignatureLineChangeImage.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("New ws");

                var stamp = ws.AddSignatureLineStamp();
                var cert = GetSelfCert();
                stamp.Sign(cert, signatureImage);

                //Arguably if there is a way to throw here we should.
                signatureImage.SetImage(GetResourceFile("EPPlus.png"));

                //Currently we throw in save
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ReadFileWithSignatureAndSignatureLineStamp()
        {
            using (var pck = OpenTemplatePackage("StampSignature.xlsx"))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets[0];

                var signature = wb.DigitialSignatures[0];
                var vmlDrawings = ws.VmlDrawings;

                var sline = ws.SignatureLines[0];
                var bytes = sline.SignatureImage.ImageBytes;

                var templateBmpFile = GetTemplateFile("StampSignatureReconstructed.bmp");

                var templateBytes = File.ReadAllBytes(templateBmpFile.FullName);
                Assert.IsTrue(bytes.SequenceEqual(templateBytes));

                SaveAndCleanup(pck);
            }
        }
    }
}
