using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Drawing;
using System.IO;

namespace EPPlusTest.Drawing.DigitalSignatures
{
    [TestClass]
    public class DigitalSignatureLineTests : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            CreatePathIfNotExists(_worksheetPath + "DigitalSignatures\\");
            ExcelPackage package = OpenPackage("DigitalSignatures\\UnsignedSignatureLine.xlsx", true);
        }

        [TestMethod]
        public void CreateAndReadDefaultSignatureLine()
        {
            var wsName = "SignatureLineWorksheet";

            using (ExcelPackage package = OpenPackage("DigitalSignatures\\UnsignedSignatureLine.xlsx"))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add(wsName);

                var sline = ws.AddSignatureLine();

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage("DigitalSignatures\\UnsignedSignatureLine.xlsx"))
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
            var signatureImage = new ExcelImage(GetResourceFile("Code.bmp"));

            using (ExcelPackage package = OpenPackage("DigitalSignatures\\UnsignedSignatureLine.xlsx"))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add(wsName);

                var slineStamp = ws.AddSignatureLineStamp();
                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage("DigitalSignatures\\UnsignedSignatureLine.xlsx"))
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

            using (ExcelPackage package = OpenPackage("DigitalSignatures\\UnsignedSignatureLine.xlsx"))
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
                sLine3.To.Row = 9 + 7;
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
                sLine5.To.Row = 9 + 7;
                sLine5.From.Column = 10;
                sLine5.To.Column = 14;

                SaveAndCleanup(package);
            }

            using (ExcelPackage package = OpenPackage("DigitalSignatures\\UnsignedSignatureLine.xlsx"))
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

                var sigStamp = ws.SignatureLines[3];
                Assert.AreEqual(sigStamp.SignatureLineType, eSignatureLineType.Stamp);
                Assert.AreEqual(Signer, sigStamp.Signer);
                Assert.AreEqual(Title, sigStamp.Title);
                Assert.AreEqual(SigningInstructions, sigStamp.SigningInstructions);
                Assert.AreEqual(ShowSignDate, sigStamp.ShowSignDate);
                Assert.AreEqual(AllowComments, sigStamp.AllowComments);
                Assert.AreEqual(AlternativeText, sigStamp.AlternativeText);
                Assert.AreEqual("{000CD6A4-0000-0000-C000-000000000046}", sigStamp.ProvID);
            }
        }

        //[TestMethod]
        //public void CreateDigitalSignatureLineStampAndSignIt()
        //{
        //    using (ExcelPackage package = OpenPackage("DigSig_SignatureLineStamp.xlsx", true))
        //    {
        //        var wb = package.Workbook;
        //        var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

        //        //wb.Calculate();

        //        var test = package.Workbook.FullCalcOnLoad;
        //        test = false;

        //        X509Store store = new X509Store(StoreLocation.CurrentUser);
        //        store.Open(OpenFlags.ReadOnly);

        //        //var aComment = ws.Comments.Add(ws.Cells["A1"], "AText");
        //        //var picture = ws.Drawings.AddPicture("APicture", "C:\\Users\\OssianEdström\\Pictures\\TempBitmap.bmp");

        //        var sLine = ws.AddSignatureLineStamp();

        //        var sLine2 = ws.AddSignatureLineStamp();
        //        var sLine3 = ws.AddSignatureLineStamp();
        //        var sLine4 = ws.AddSignatureLineStamp();
        //        var sLine5 = ws.AddSignatureLineStamp();


        //        sLine2.From.Column = 3;
        //        sLine2.To.Column = 5;

        //        sLine3.From.Column = 5;
        //        sLine3.To.Column = 7;

        //        sLine3.Signer = "SomeOne";

        //        sLine4.From.Column = 7;
        //        sLine4.To.Column = 9;

        //        sLine5.From.Column = 9;
        //        sLine5.To.Column = 11;

        //        //sLine2.From.Row = 1;
        //        //sLine2.To.Row = 8;

        //        //sLine3.From.ColumnOffset = 200;
        //        //sLine4.From.ColumnOffset = 300;

        //        sLine.Signer = "ASigner";
        //        sLine.Title = "Developer";
        //        //sLine.IsStamp = true;
        //        sLine.SignatureImage = new ExcelImage("C:\\Users\\OssianEdström\\Pictures\\TempBitmap.bmp");
        //        //sLine.SignatureImage = File.ReadAllBytes("C:\\Users\\OssianEdström\\Pictures\\ghostlady.png");

        //        var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1], CommitmentType.CreatedAndApproved, "TestingSignatureLine");
        //        var info = digSig.SigningInformation;

        //        info.SignerRoleTitle = "A Title";
        //        info.Address1 = "Some";
        //        info.Address2 = "Where";
        //        info.ZIPorPostalCode = "Over";
        //        info.City = "The";
        //        info.CountryOrRegion = "Rainbow";
        //        info.StateOrProvince = "WayUpHigh";

        //        digSig.SignatureLine = sLine;

        //        var digSig2 = wb.DigitialSignatures.AddSignature(store.Certificates[0], CommitmentType.Created, "TestingSignatureLine2");
        //        digSig2.SignatureLine = sLine2;

        //        SaveAndCleanup(package);
        //    }
        //}

        [TestMethod]
        public void FixGeneratedImage()
        {
            var generated = new EmfImage();
            generated.Read("C:\\epplusTest\\Testoutput\\image1Generated.emf");
            var records = generated.records;

            var clipRect = (EMR_INTERSECTCLIPRECT)records[58];
            clipRect.Clip = new RectLObject(35,4,93,17);
            generated.Save("C:\\epplusTest\\Testoutput\\generatedImageRechanged.emf");
        }

        [TestMethod]
        public void testImage()
        {
            var generated = new EmfImage();
            generated.Read(@"C:\epplusTest\templates\maxedBars.emf");
            var records = generated.records;

            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Cast<EMR_EXTTEXTOUTW>().ToList();

            textRecords[0].Text = "MMMMMMMMMM||";
            textRecords[0].AdjustReferenceToCenterText(127, 10);

            generated.Save("C:\\epplusTest\\Testoutput\\MaxTitleGenned.emf");
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
        public void TemplateTest()
        {
            var stampTemplate = new EmfImage();
            stampTemplate.Read(@"C:\epplusTest\templates\SignatureLineStampTemplate.emf");

            var records = stampTemplate.records;

            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);

            ((EMR_EXTTEXTOUTW)textRecords[0]).Text = "";
            //((EMR_EXTTEXTOUTW)textRecords[1]).Text = "";

            stampTemplate.Save(@"C:\epplusTest\templates\SignatureLineStampTemplateNew.emf");
        }

        [TestMethod]
        public void CreateTwoEmptySignatureLine()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_Empty2.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();

                var sLine2 = ws.AddSignatureLine();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CreateSignatureLineWithSuggestedSigner()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_SSigner.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "ASuggestedSigner";

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void CreateSignatureLineWithSuggestedSignerAndTitle()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_SSignerTitle.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "ASuggestedSigner";
                sLine.Title = "ASuggestedTitle";

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CreateSignatureLineWithALL()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_ALL.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "ASuggestedSigner";
                sLine.Title = "ASuggestedTitle";
                sLine.Email = "Example@Site.com";
                sLine.SigningInstructions = "Hey please sign this because x and y so it will be z";
                sLine.AllowComments = true;
                sLine.ShowSignDate = true;

                SaveAndCleanup(package);
            }
        }
    }
}
