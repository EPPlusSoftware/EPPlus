﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security.Cryptography;
using System.Xml;
using System.Security.Cryptography.Xml;
using System.IO;
using System;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using OfficeOpenXml.DigitalSignatures;
using System.Collections.Generic;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Drawing.EMF;
using System.Linq;

//REMEMBER:
//1. Cannonize
//2. Transform
//3. Hash data

namespace EPPlusTest.Drawing.DigitalSignatures
{
    [TestClass]
    public class DigitalSignatureTests : TestBase
    {
        [TestMethod]
        public void CreateDigitalSignatureLine()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                wb.Calculate();

                var test = package.Workbook.FullCalcOnLoad;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);

                //var sigLine = ws.Workbook.DigitialSignatures.AddSignatureLine(store.Certificates[1], ws);
                //sigLine.VmlDrawing.Signer = "AttemptedSigner";
                //sigLine.VmlDrawing.Title = "AttemptedTitle";

                var digSig = ws.Workbook.DigitialSignatures.AddSignature(store.Certificates[1], CommitmentType.CreatedAndApproved, "TestingSignatureLine");
                var info = digSig.SignerInformation;

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZIPorPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void EncodeEmf()
        {
            var fileName = "C:\\epplusTest\\Testoutput\\image1.emf";
            var invalidLnImg = Convert.ToBase64String(File.ReadAllBytes(fileName));
            Assert.AreEqual(invalidLnImg, "AQAAAGwAAAAAAAAAAAAAAP8AAAB/AAAAAAAAAAAAAABcFwAAqwsAACBFTUYAAAEAsB8AALEAAAAGAAAAAAAAAAAAAAAAAAAAAAoAAKAFAABWAgAAUAEAAAAAAAAAAAAAAAAAAPAfCQCAIAUACgAAABAAAAAAAAAAAAAAAEsAAAAQAAAAAAAAAAUAAAAeAAAAGAAAAAAAAAAAAAAAAAEAAIAAAAAnAAAAGAAAAAEAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAEAAABMAAAAZAAAAAAAAAAAAAAA/wAAAH8AAAAAAAAAAAAAAAABAACAAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAABAAAAJwAAABgAAAABAAAAAAAAAP///wAAAAAAJQAAAAwAAAABAAAATAAAAGQAAAAAAAAAAAAAAP8AAAB/AAAAAAAAAAAAAAAAAQAAgAAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAQAAACcAAAAYAAAAAQAAAAAAAADw8PAAAAAAACUAAAAMAAAAAQAAAEwAAABkAAAAAAAAAAAAAAD/AAAAfwAAAAAAAAAAAAAAAAEAAIAAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAEAAAAnAAAAGAAAAAEAAAAAAAAA8PDwAAAAAAAlAAAADAAAAAEAAABMAAAAZAAAAAAAAAAAAAAA/wAAAH8AAAAAAAAAAAAAAAABAACAAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAABAAAAJwAAABgAAAABAAAAAAAAAPDw8AAAAAAAJQAAAAwAAAABAAAATAAAAGQAAAAAAAAAAAAAAP8AAAB/AAAAAAAAAAAAAAAAAQAAgAAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAQAAACcAAAAYAAAAAQAAAAAAAADw8PAAAAAAACUAAAAMAAAAAQAAAEwAAABkAAAAAAAAAAAAAAD/AAAAfwAAAAAAAAAAAAAAAAEAAIAAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAEAAAAnAAAAGAAAAAEAAAAAAAAA////AAAAAAAlAAAADAAAAAEAAABMAAAAZAAAAAAAAAAAAAAA/wAAAH8AAAAAAAAAAAAAAAABAACAAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAABAAAAJwAAABgAAAABAAAAAAAAAP///wAAAAAAJQAAAAwAAAABAAAATAAAAGQAAAAAAAAAAAAAAP8AAAB/AAAAAAAAAAAAAAAAAQAAgAAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAQAAACcAAAAYAAAAAQAAAAAAAAD///8AAAAAACUAAAAMAAAAAQAAAEwAAABkAAAAAAAAAAMAAAD/AAAAEgAAAAAAAAADAAAAAAEAABAAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAEAAAAnAAAAGAAAAAEAAAAAAAAA////AAAAAAAlAAAADAAAAAEAAABMAAAAZAAAAAkAAAADAAAAGAAAABIAAAAJAAAAAwAAABAAAAAQAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAABAAAAFQAAAAwAAAADAAAAcgAAALADAAAKAAAAAwAAABcAAAAQAAAA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AQICIlMAeQBzAHQAZQBtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZHYACAAAAAAlAAAADAAAAAIAAAAnAAAAGAAAAAMAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAMAAABMAAAAZAAAAAAAAAAAAAAA//////////8AAAAAFgAAAAAAAAA1AAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAADAAAAJwAAABgAAAADAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAADAAAATAAAAGQAAAAAAAAAAAAAAP//////////AAAAABYAAAAAAQAAAAAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAwAAACcAAAAYAAAAAwAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAwAAAEwAAABkAAAAAAAAAAAAAAD//////////wABAAAWAAAAAAAAADUAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAMAAAAnAAAAGAAAAAMAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAMAAABMAAAAZAAAAAAAAABLAAAA/wAAAEwAAAAAAAAASwAAAAABAAACAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAADAAAAJwAAABgAAAADAAAAAAAAAP///wAAAAAAJQAAAAwAAAADAAAATAAAAGQAAAAAAAAAFgAAAP8AAABKAAAAAAAAABYAAAAAAQAANQAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAAAwAAACcAAAAYAAAAAwAAAAAAAAD///8AAAAAACUAAAAMAAAAAwAAAEwAAABkAAAACQAAACcAAAAfAAAASgAAAAkAAAAnAAAAFwAAACQAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAMAAABSAAAAcAEAAAMAAADg////AAAAAAAAAAAAAAAAkAEAAAAAAAEAAAAAYQByAGkAYQBsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkdgAIAAAAACUAAAAMAAAAAwAAABgAAAAMAAAAAAAAABIAAAAMAAAAAQAAABYAAAAMAAAACAAAAFQAAABUAAAACgAAACcAAAAeAAAASgAAAAEAAAAA4LpBq6q6QQoAAABLAAAAAQAAAEwAAAAEAAAACQAAACcAAAAgAAAASwAAAFAAAABYAAAAFQAAABYAAAAMAAAAAAAAACUAAAAMAAAAAgAAACcAAAAYAAAABAAAAAAAAAD///8AAAAAACUAAAAMAAAABAAAAEwAAABkAAAAKQAAABkAAAD2AAAASgAAACkAAAAZAAAAzgAAADIAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAAAAIAoAAAADAAAAAQAAAAnAAAAGAAAAAQAAAAAAAAA////AAAAAAAlAAAADAAAAAQAAABMAAAAZAAAACkAAAAZAAAA9gAAAEcAAAApAAAAGQAAAM4AAAAvAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAAAACAKAAAAAwAAAAEAAAAJwAAABgAAAAEAAAAAAAAAP///wAAAAAAJQAAAAwAAAAEAAAATAAAAGQAAAApAAAAMwAAAFkAAABHAAAAKQAAADMAAAAxAAAAFQAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAAAAgCgAAAAMAAAABAAAAFIAAABwAQAABAAAAPD///8AAAAAAAAAAAAAAACQAQAAAAAAAQAAAABzAGUAZwBvAGUAIAB1AGkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGR2AAgAAAAAJQAAAAwAAAAEAAAAGAAAAAwAAAAAAAAA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");
            var bytes = HashAndEncodeBytes(File.ReadAllBytes("C:\\epplusTest\\Testoutput\\image1.emf"));
        }

        private void DecodeAndSaveEmf(string base64String, string savePath)
        {
            var decodedBytes = Convert.FromBase64String(base64String);
            File.WriteAllBytes(savePath, decodedBytes);
        }


        [TestMethod]
        public void ReadEmfSpacing()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\ValidImageAlternateSize.emf");
        }

        [TestMethod]
        public void CheckValidTemplate()
        {
            var validTemplate = new SignatureLineTemplateValid();
            var records = validTemplate.records;

            validTemplate.timeStamp.Text = "TimeStamp";
            validTemplate.signTextObject.Text = "TemplateSignature";
            validTemplate.suggestedSignerObject.Text = "TemplateSigner";
            validTemplate.suggestedTitleObject.Text = "TemplateTitle";
            validTemplate.SignedBy = "TemplateName";

            validTemplate.Save("C:\\epplusTest\\Testoutput\\testTemp.emf");
        }

        [TestMethod]
        public void TestTextLength()
        {
            var inValidTemplate = new SignatureLineTemplateInvalid();
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
        public void CheckInValidTemplate()
        {
            var inValidTemplate = new SignatureLineTemplateInvalid();
            var records = inValidTemplate.records;

            inValidTemplate.SignText = "IHaveAVeryVeryVeryVerylon";
            inValidTemplate.suggestedSignerObject.Text = "TemplateSigner";
            inValidTemplate.suggestedTitleObject.Text = "TemplateTitle";
            inValidTemplate.SignedBy = "TemplateName";

            inValidTemplate.Save("C:\\epplusTest\\Testoutput\\TempTest.emf");
        }

        [TestMethod]
        public void ReadEmf()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\LongName.emf");

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(2);
            var arr = textRecordArr.ToArray();

            var longName = (EMR_EXTTEXTOUTW)arr[0];
            var suggestedSigner = (EMR_EXTTEXTOUTW)arr[1];

            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);

            var longIndex = emfImage.records.IndexOf(longName);
            var signerIndex = emfImage.records.IndexOf(suggestedSigner);

            emfImage.records[140].data = new byte[] { 3, 0, 0, 0 };

            emfImage.Save("C:\\epplusTest\\Testoutput\\ChangeFontOutput.emf");


            //for (int i = 0; i < textRecordArr.Count(); i++)
            //{
            //    var textRecord = (EMR_EXTTEXTOUTW)arr[i];
            //    textRecord.Text = TemplateNamesArr[i];
            //}

            //var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);

            //var textRecordArrTest = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);
            //var arrTest = textRecordArrTest.ToArray();

            //((EMR_EXTTEXTOUTW)arrTest[1]).Text = "Y";

            //var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(2);
            //var arr  = textRecordArr.ToArray();

            //var imageRecord = emfImage.records.IndexOf(arr[0]);

            //var TemplateNamesArr = new string[] { "TemplateSignature", "TemplateSigner", "TemplateTitle", "Signed by: TemplateName" };

            ////((EMR_EXTTEXTOUTW)arr[0]).Text= "Tea";
            ////((EMR_EXTTEXTOUTW)arr[1]).Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            ////((EMR_EXTTEXTOUTW)arr[2]).Text = "Test2";
            ////((EMR_EXTTEXTOUTW)arr[3]).Text = "Test3";

            //for (int i = 0; i < textRecordArr.Count(); i++)
            //{
            //    var textRecord = (EMR_EXTTEXTOUTW)arr[i];
            //    textRecord.Text = TemplateNamesArr[i];
            //}

            ////foreach (var record in textRecordArr)
            ////{
            ////    var txtRecord = ((EMR_EXTTEXTOUTW)record);
            ////    txtRecord.Text = "templateText";
            ////}

            //emfImage.Save("C:\\epplusTest\\Testoutput\\InvalidSignatureLineTemplate2.emf");
        }

        [TestMethod]
        public void SavingEmptyPartShouldCreateFileAndNotThrow()
        {
            using (ExcelPackage package = new ExcelPackage("DigSig_EmptyPart"))
            {
                package.Workbook.Worksheets.Add("newWorksheet");
                string partURI = @"/_xmlsignatures/origin.sigs";
                var part = package.ZipPackage.CreatePart(new Uri(partURI, UriKind.Relative), ContentTypes.signatureOrigin);
                var stream = part.GetStream();
                stream.Write([], 0, 0);
                part.CreateRelationship("sig1.xml", TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature");

                SaveAndCleanup(package);
            }

            //TODO: Read file and verify existance
        }

        [TestMethod]
        public void VerifyTheory2()
        {
            //CspParameters cspParams = new()
            //{
            //    KeyContainerName = "XML_DSIG_RSA_KEY",
            //};

            // RSACryptoServiceProvider rsaKey = new(cspParams);

            RSACryptoServiceProvider rsaKey = new();

            XmlDocument xmlDoc = new()
            {
                PreserveWhitespace = true,
            };
            xmlDoc.Load("C:\\epplusTest\\Workbooks\\idOfficeSeparateNew.xml");

            SignedXml signedXml = new(xmlDoc)
            {
                SigningKey = rsaKey
            };

            Reference reference = new()
            {
                Type = "http://www.w3.org/2000/09/xmldsig#Object",
                Uri = "#idOfficeObject"
            };

            reference.DigestMethod = "http://www.w3.org/2000/09/xmldsig#sha1";

            var idElement = signedXml.GetIdElement(xmlDoc, "idOfficeObject");

            signedXml.Signature.Id = "idPackageSignature";

            //Verify file from earlier ------------
            XmlDocument earlyXmlDoc = new()
            {
                PreserveWhitespace = true,
            };
            earlyXmlDoc.Load("C:\\epplusTest\\Workbooks\\sig1TestFile.xml");
            SignedXml signedXmlEarly = new(earlyXmlDoc)
            {
                SigningKey = rsaKey
            };

            var earlyIdElement = signedXmlEarly.GetIdElement(earlyXmlDoc, "idOfficeObject");

            Assert.AreEqual(earlyIdElement.OuterXml, idElement.OuterXml);
            //------End of verify-----------

            signedXml.AddReference(reference);

            signedXml.ComputeSignature();

            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            XmlElement xmlDigitalSignature = signedXml.GetXml();

            XmlDocument doc = new XmlDocument()
            {
                PreserveWhitespace = true,
            };

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            doc.DocumentElement?.AppendChild(doc.ImportNode(xmlDigitalSignature, true));

            doc.Save("C:\\epplusTest\\Workbooks\\newVersion.xml");
        }

        internal const string PartUri = @"/_xmlsignatures/sig1.xml";

        [TestMethod]
        public void streamTest()
        {
            //var stream = new MemoryStream();
            //stream.
            //stream.Write([], 0, 0);
            //File.WriteAllBytes("C:\\Users\\OssianEdström\\Documents\\AlignmentTest-OnCells.xlsx", stream.get);
        }

        [TestMethod]
        public void SignAsExcelDoes()
        {

            //CspParameters cspParams = new()
            //{
            //    KeyContainerName = "XML_DSIG_RSA_KEY",
            //};

            // RSACryptoServiceProvider rsaKey = new(cspParams);

            RSACryptoServiceProvider rsaKey = new();

            XmlDocument xmlDoc = new()
            {
                PreserveWhitespace = true,
            };

            xmlDoc.Load("C:\\epplusTest\\Workbooks\\sig1TestFile.xml");

            SignedXml signedXml = new(xmlDoc)
            {
                SigningKey = rsaKey
            };

            Reference reference = new()
            {
                Type = "http://www.w3.org/2000/09/xmldsig#Object",
                Uri = "#idOfficeObject"
            };

            reference.DigestMethod = "http://www.w3.org/2000/09/xmldsig#sha1";

            var idElement = signedXml.GetIdElement(xmlDoc, "idOfficeObject");

            XmlDocument xmlDoc3 = new()
            {
                PreserveWhitespace = true,
            };
            xmlDoc3.Load("C:\\epplusTest\\Workbooks\\newXml.xml");


            signedXml.Signature.Id = "idPackageSignature";

            XmlDocument doc = new XmlDocument()
            {
                PreserveWhitespace = true,
            };

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //Breaking out idOffice part of file to new file for later verification.
            var aString = idElement.OuterXml;
            File.WriteAllText("C:\\epplusTest\\Workbooks\\idOfficeSeparateNew.xml", aString, Encoding.UTF8);

            signedXml.AddReference(reference);

            signedXml.ComputeSignature();

            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            XmlElement xmlDigitalSignature = signedXml.GetXml();

            XmlDocument xmlDoc2 = new()
            {
                PreserveWhitespace = true,
            };
            xmlDoc2.Load("C:\\epplusTest\\Workbooks\\newXml.xml");

            xmlDoc2.DocumentElement?.AppendChild(xmlDoc2.ImportNode(xmlDigitalSignature, true));

            var listNodes = xmlDigitalSignature.GetElementsByTagName("DigestValue");
            var node1 = listNodes[0];

            Assert.AreEqual("Dwx/mtIT+lffP980qEOPVRJX41k=", node1.InnerText);
            //var stringTest = System.Text.Encoding.UTF8.GetString(reference.DigestValue);
            //Assert.AreEqual("Dwx/mtIT+lffP980qEOPVRJX41k=", stringTest);

            xmlDoc2.Save("C:\\epplusTest\\Workbooks\\newVersionExcelBased.xml");
        }

        [TestMethod]
        public void SignSave()
        {
            using (var pck = OpenPackage("generatedSignedEmpty.xlsx", true))
            {
                RSACryptoServiceProvider rsaKey = new();

                var wb = pck.Workbook;

                var ws = wb.Worksheets.Add("emptyWorksheet");

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1]);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void VerifyTheoryIndividualFiles()
        {
            var styles = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\genStyles.xml");
            var hashedStyles = HashAndEncodeBytes(styles);
            Assert.AreEqual("R3jSMFWoLJ87ma2wdBoixK+0JNU=", hashedStyles);
        }

        [TestMethod]
        public void VerifyTheoryIndividualFiles2()
        {
            var styles = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\stylesUnsigned.xml");
            MemoryStream newMs = RecyclableMemory.GetStream();
            newMs.Write(styles, 0, styles.Length);
            newMs.Position = 0;
            string EndString;

            using (StreamReader reader = new StreamReader(newMs))
            {
                EndString = reader.ReadToEnd();
            }

            var hashedStyles = HashAndEncodeBytes(styles);
            Assert.AreEqual("R3jSMFWoLJ87ma2wdBoixK+0JNU=", hashedStyles);
        }

        [TestMethod]
        public void VerifyTheoryIndividualFiles3()
        {
            var styles = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\epplusGennedStyleComplete.xml");
            var hashedStyles = HashAndEncodeBytes(styles);
            Assert.AreEqual("R3jSMFWoLJ87ma2wdBoixK+0JNU=", hashedStyles);
        }

        [TestMethod]
        public void SignRelationshipFile()
        {
            //XmlDocument document = new XmlDocument();

            var bytes = File.ReadAllBytes("C:\\Users\\OssianEdström\\Documents\\.rels");

            MemoryStream stream = new MemoryStream(bytes);

            RSACryptoServiceProvider rsaKey = new();

            SignedXml signedXml = new()
            {
                SigningKey = rsaKey,
            };

            signedXml.SignedInfo.CanonicalizationMethod = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            Reference _ref = new(stream);

            var commentsTransform = new XmlDsigC14NTransform();

            _ref.AddTransform(commentsTransform);
            _ref.DigestMethod = "http://www.w3.org/2000/09/xmldsig#sha1";

            signedXml.AddReference(_ref);
            signedXml.ComputeSignature();

            var test = signedXml.GetXml().OuterXml;
        }

        [TestMethod]
        public void SignExact()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("C:\\epplusTest\\Personal Compare\\Test\\objectsOnly.xml");

            //RSA key = RSA.Create();

            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            var Certificate = store.Certificates[0];

            var x509KeyInfo = new KeyInfoX509Data(Certificate);

            var rsaKey = Certificate.GetRSAPrivateKey();

            ExcelSignedXml signedXml = new(doc)
            {
                SigningKey = rsaKey,
            };

            signedXml.Signature.Id = "idPackageSignature";
            signedXml.SignedInfo.CanonicalizationMethod = SignedXml.XmlDsigCanonicalizationUrl;
            signedXml.SignedInfo.SignatureMethod = SignedXml.XmlDsigRSASHA1Url;

            KeyInfo keyInfo = new KeyInfo();

            KeyInfoX509Data clause = new KeyInfoX509Data();
            //clause.AddSubjectName(Certificate.Subject);
            clause.AddCertificate(Certificate);
            keyInfo.AddClause(clause);
            signedXml.KeyInfo = keyInfo;


            Reference packageReference = new()
            {
                Type = "http://www.w3.org/2000/09/xmldsig#Object",
                Uri = "#idPackageObject"
            };

            packageReference.DigestMethod = DigestMethods.SHA1;

            var packageObj = new DataObject();
            var pElement = (XmlElement)doc.GetElementsByTagName("Object")[0];
            packageObj.LoadXml(pElement);

            signedXml.AddObject(packageObj);
            signedXml.AddReference(packageReference);

            Reference OfficeReference = new()
            {
                Type = "http://www.w3.org/2000/09/xmldsig#Object",
                Uri = "#idOfficeObject"
            };

            OfficeReference.DigestMethod = DigestMethods.SHA1;

            var officeObj = new DataObject();
            var oElement = (XmlElement)doc.GetElementsByTagName("Object")[1];
            officeObj.LoadXml(oElement);

            signedXml.AddObject(officeObj);
            signedXml.AddReference(OfficeReference);

            Reference signedPropertiesReference = new()
            {
                Type = "http://uri.etsi.org/01903#SignedProperties",
                Uri = "#idSignedProperties"
            };
            XmlDsigC14NTransform c14Transform = new();

            signedPropertiesReference.AddTransform(c14Transform);
            signedPropertiesReference.DigestMethod = DigestMethods.SHA1;

            DataObject signedProps = new DataObject();
            var sElement = (XmlElement)doc.GetElementsByTagName("Object")[2];
            signedProps.LoadXml(sElement);

            signedXml.AddObject(signedProps);
            signedXml.AddReference(signedPropertiesReference);

            signedXml.ComputeSignature();

            XmlElement xmlDigitalSignature = signedXml.GetXml();

            var output = new XmlDocument() { PreserveWhitespace = true };

            var node = output.ImportNode(xmlDigitalSignature, true);
            output.AppendChild(node);

            var declaration = output.CreateXmlDeclaration("1.0", "UTF-8", "");
            output.InsertBefore(declaration, node);

            var check = signedXml.CheckSignature(rsaKey);

            var defaultValue = Convert.ToBase64String(signedXml.Signature.SignatureValue);
            var original = Convert.ToBase64String(signedXml.Signature.SignatureValue, Base64FormattingOptions.InsertLineBreaks);

            var sigValueElement = output.GetElementsByTagName("SignatureValue")[0];

            sigValueElement.InnerText = Convert.ToBase64String(signedXml.SignatureValue, Base64FormattingOptions.InsertLineBreaks);

            var arr = signedXml.Signature.SignatureValue;
            Array.Reverse(arr);

            var reversed = Convert.ToBase64String(arr);
            var reversedWithLineBreaks = Convert.ToBase64String(signedXml.Signature.SignatureValue, Base64FormattingOptions.InsertLineBreaks);

            output.Save("C:\\epplusTest\\Personal Compare\\Test\\epplusOutput.xml");
        }

        [TestMethod]
        public void RelTest()
        {
            var bytes = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\.rels");
            XmlDocument document = new XmlDocument();
            document.LoadXml(Encoding.UTF8.GetString(bytes));
            var relationTransform = new RelTransform(document, new List<string> { "rId1" });

            var someXml = relationTransform.GetOutputXML();

            var newBytes = Encoding.Default.GetBytes(someXml);

            MemoryStream stream = new MemoryStream(newBytes);

            RSACryptoServiceProvider rsaKey = new();

            SignedXml signedXml = new()
            {
                SigningKey = rsaKey,
            };

            signedXml.SignedInfo.CanonicalizationMethod = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            Reference _ref = new(stream);

            var commentsTransform = new XmlDsigC14NTransform();

            _ref.AddTransform(commentsTransform);
            _ref.DigestMethod = "http://www.w3.org/2000/09/xmldsig#sha1";

            signedXml.AddReference(_ref);
            signedXml.ComputeSignature();

            var test = signedXml.GetXml();

            XmlDocument xmlDocument = new XmlDocument();

            xmlDocument.LoadXml(test.OuterXml);

            var transforms = xmlDocument.GetElementsByTagName("Transforms");

            var relElement = new XmlDocument();
            relElement.LoadXml(relationTransform.TransformXml);

            var importedNode = xmlDocument.ImportNode(relElement.FirstChild, true);

            transforms[0].InsertBefore(importedNode, transforms[0].FirstChild);

            var strTest = xmlDocument.GetElementsByTagName("Reference")[0].OuterXml;
        }


        [TestMethod]
        public void TestEncode()
        {
            var issueText = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>";

            var textBytes = Encoding.UTF8.GetBytes(issueText);

            var hashed = HashAndEncodeBytes(textBytes);
        }

        [TestMethod]
        public void TestEncodeFile()
        {
            var textBytes = File.ReadAllBytes("C:\\epplusTest\\Personal Compare\\ResignedExcel.xlsx\\xl\\sharedStrings.xml");
            //var issueText = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>";

            //var textBytes = Encoding.UTF8.GetBytes(issueText)
            //;
            Assert.AreEqual(
            new byte[] { textBytes[0], textBytes[1], textBytes[2] },
            Encoding.UTF8.GetPreamble());

            var hashed = HashAndEncodeBytes(textBytes);
        }

        [TestMethod]
        public void HashWithJustTransform()
        {
            XmlDocument readDoc = new XmlDocument() { PreserveWhitespace = true };
            readDoc.Load("C:\\epplusTest\\Personal Compare\\signedProperties2.xml");

            XmlElement element = (XmlElement)readDoc.GetElementsByTagName("xd:SignedProperties")[0];
            element.SetAttribute("xmlns", "http://www.w3.org/2000/09/xmldsig#");

            XmlDocument signatureDocument = new XmlDocument() { PreserveWhitespace = true };

            var sourceNode = signatureDocument.ImportNode(element, true);
            signatureDocument.AppendChild(sourceNode);

            var testNodes = signatureDocument.SelectNodes("//*[name()=local-name()]");

            foreach (XmlElement node in testNodes)
            {
                node.SetAttribute("xmlns", "http://www.w3.org/2000/09/xmldsig#");
            }

            var bytes = Encoding.UTF8.GetBytes(signatureDocument.OuterXml);

            var streamTest = new MemoryStream(bytes);

            XmlDsigC14NTransform testTransform = new();

            testTransform.LoadInput(streamTest);

            var digested = testTransform.GetDigestedOutput(SHA1.Create());
            var stringedDigest = Convert.ToBase64String(digested);

            Assert.AreEqual("n8llW6rkqfPAu2g024cwGvHKS3Y=", stringedDigest);
        }

        [TestMethod]
        public void HashQualifyingPropertiesCorrectly()
        {
            XmlDocument xmlDocument = new XmlDocument() { PreserveWhitespace = true };

            var root = xmlDocument.CreateElement("Signature", "http://www.w3.org/2000/09/xmldsig#");
            xmlDocument.AppendChild(root);

            XmlDocument readDoc = new XmlDocument() { PreserveWhitespace = true };
            //readDoc.Load("C:\\epplusTest\\Workbooks\\signedProperties.xml");
            readDoc.Load("C:\\epplusTest\\Personal Compare\\signedProperties2.xml");

            readDoc.DocumentElement.SetAttribute("xmlns", "http://www.w3.org/2000/09/xmldsig#");
            var test = readDoc.DocumentElement.InnerXml;

            root.InnerXml = readDoc.InnerXml;

            xmlDocument.DocumentElement.ChildNodes[0].Attributes.RemoveNamedItem("xmlns");

            RSACryptoServiceProvider rsaKey = new();

            ExcelSignedXml signedXml = new(xmlDocument)
            {
                SigningKey = rsaKey,
            };

            signedXml.Signature.Id = "idPackageSignature";
            signedXml.SignedInfo.CanonicalizationMethod = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            var dO = new DataObject();
            dO.Data = xmlDocument.DocumentElement.ChildNodes;

            signedXml.AddObject(dO);

            Reference signedPropertiesReference = new()
            {
                Type = "http://uri.etsi.org/01903#SignedProperties",
                Uri = "#idSignedProperties"
            };
            XmlDsigC14NTransform c14Transform = new();

            var elementDoc = new XmlDocument() { PreserveWhitespace = true };
            var idElement = signedXml.GetIdElement(xmlDocument, "idSignedProperties");
            var idElementImported = elementDoc.ImportNode(idElement, true);
            elementDoc.AppendChild(idElementImported);

            signedPropertiesReference.AddTransform(c14Transform);
            signedPropertiesReference.DigestMethod = DigestMethods.SHA1;

            signedXml.AddReference(signedPropertiesReference);

            signedXml.ComputeSignature();

            var xml = signedXml.GetXml();

            XmlDsigC14NTransform testTransform = new();

            var digested = testTransform.GetDigestedOutput(SHA1.Create());
            var stringedDigest = Convert.ToBase64String(digested);
        }

        //Normalize
        //Canonize
        //Transform
        //Hash
        //Read as string
        [TestMethod]
        public void PurelyTransform()
        {
            var propertiesRaw = "<Object><xd:QualifyingProperties xmlns:xd=\"http://uri.etsi.org/01903/v1.3.2#\" Target=\"#idPackageSignature\"><xd:SignedProperties Id=\"idSignedProperties\"><xd:SignedSignatureProperties><xd:SigningTime>2024-08-14T07:37:36Z</xd:SigningTime><xd:SigningCertificate><xd:Cert><xd:CertDigest><DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/><DigestValue>w9iTMIvTXcdRc9G38Pp1Njb/HPE=</DigestValue></xd:CertDigest><xd:IssuerSerial><X509IssuerName>CN=OssianEdström</X509IssuerName><X509SerialNumber>38225183535545048482234589307877617536</X509SerialNumber></xd:IssuerSerial></xd:Cert></xd:SigningCertificate><xd:SignaturePolicyIdentifier><xd:SignaturePolicyImplied/></xd:SignaturePolicyIdentifier></xd:SignedSignatureProperties><xd:SignedDataObjectProperties><xd:CommitmentTypeIndication><xd:CommitmentTypeId><xd:Identifier>http://uri.etsi.org/01903/v1.2.2#ProofOfOrigin</xd:Identifier><xd:Description>Created and approved this document</xd:Description></xd:CommitmentTypeId><xd:AllSignedDataObjects/><xd:CommitmentTypeQualifiers><xd:CommitmentTypeQualifier>Forty-Two</xd:CommitmentTypeQualifier></xd:CommitmentTypeQualifiers></xd:CommitmentTypeIndication></xd:SignedDataObjectProperties></xd:SignedProperties></xd:QualifyingProperties></Object>";
            var propertiesFromId = "<xd:SignedProperties Id=\"idSignedProperties\"><xd:SignedSignatureProperties><xd:SigningTime>2024-08-14T07:37:36Z</xd:SigningTime><xd:SigningCertificate><xd:Cert><xd:CertDigest><DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/><DigestValue>w9iTMIvTXcdRc9G38Pp1Njb/HPE=</DigestValue></xd:CertDigest><xd:IssuerSerial><X509IssuerName>CN=OssianEdström</X509IssuerName><X509SerialNumber>38225183535545048482234589307877617536</X509SerialNumber></xd:IssuerSerial></xd:Cert></xd:SigningCertificate><xd:SignaturePolicyIdentifier><xd:SignaturePolicyImplied/></xd:SignaturePolicyIdentifier></xd:SignedSignatureProperties><xd:SignedDataObjectProperties><xd:CommitmentTypeIndication><xd:CommitmentTypeId><xd:Identifier>http://uri.etsi.org/01903/v1.2.2#ProofOfOrigin</xd:Identifier><xd:Description>Created and approved this document</xd:Description></xd:CommitmentTypeId><xd:AllSignedDataObjects/><xd:CommitmentTypeQualifiers><xd:CommitmentTypeQualifier>Forty-Two</xd:CommitmentTypeQualifier></xd:CommitmentTypeQualifiers></xd:CommitmentTypeIndication></xd:SignedDataObjectProperties></xd:SignedProperties>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(propertiesRaw);

            XmlDsigC14NTransform aTransform = new();

            var signedPropElement = doc.GetElementsByTagName("xd:SignedProperties");
            aTransform.LoadInput(doc);

            MemoryStream stream = aTransform.GetOutput() as MemoryStream;

            var digested = aTransform.GetDigestedOutput(SHA1.Create());
            var stringedDigest = Convert.ToBase64String(digested);

        }

        [TestMethod]
        public void HashingQualifyingPropertiesTest()
        {
            var propertiesRaw = "<Object><xd:QualifyingProperties xmlns:xd=\"http://uri.etsi.org/01903/v1.3.2#\" Target=\"#idPackageSignature\"><xd:SignedProperties Id=\"idSignedProperties\"><xd:SignedSignatureProperties><xd:SigningTime>2024-08-14T07:37:36Z</xd:SigningTime><xd:SigningCertificate><xd:Cert><xd:CertDigest><DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/><DigestValue>w9iTMIvTXcdRc9G38Pp1Njb/HPE=</DigestValue></xd:CertDigest><xd:IssuerSerial><X509IssuerName>CN=OssianEdström</X509IssuerName><X509SerialNumber>38225183535545048482234589307877617536</X509SerialNumber></xd:IssuerSerial></xd:Cert></xd:SigningCertificate><xd:SignaturePolicyIdentifier><xd:SignaturePolicyImplied/></xd:SignaturePolicyIdentifier></xd:SignedSignatureProperties><xd:SignedDataObjectProperties><xd:CommitmentTypeIndication><xd:CommitmentTypeId><xd:Identifier>http://uri.etsi.org/01903/v1.2.2#ProofOfOrigin</xd:Identifier><xd:Description>Created and approved this document</xd:Description></xd:CommitmentTypeId><xd:AllSignedDataObjects/><xd:CommitmentTypeQualifiers><xd:CommitmentTypeQualifier>Forty-Two</xd:CommitmentTypeQualifier></xd:CommitmentTypeQualifiers></xd:CommitmentTypeIndication></xd:SignedDataObjectProperties></xd:SignedProperties></xd:QualifyingProperties></Object>";
            var propertiesFromId = "<xd:SignedProperties Id=\"idSignedProperties\"><xd:SignedSignatureProperties><xd:SigningTime>2024-08-14T07:37:36Z</xd:SigningTime><xd:SigningCertificate><xd:Cert><xd:CertDigest><DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/><DigestValue>w9iTMIvTXcdRc9G38Pp1Njb/HPE=</DigestValue></xd:CertDigest><xd:IssuerSerial><X509IssuerName>CN=OssianEdström</X509IssuerName><X509SerialNumber>38225183535545048482234589307877617536</X509SerialNumber></xd:IssuerSerial></xd:Cert></xd:SigningCertificate><xd:SignaturePolicyIdentifier><xd:SignaturePolicyImplied/></xd:SignaturePolicyIdentifier></xd:SignedSignatureProperties><xd:SignedDataObjectProperties><xd:CommitmentTypeIndication><xd:CommitmentTypeId><xd:Identifier>http://uri.etsi.org/01903/v1.2.2#ProofOfOrigin</xd:Identifier><xd:Description>Created and approved this document</xd:Description></xd:CommitmentTypeId><xd:AllSignedDataObjects/><xd:CommitmentTypeQualifiers><xd:CommitmentTypeQualifier>Forty-Two</xd:CommitmentTypeQualifier></xd:CommitmentTypeQualifiers></xd:CommitmentTypeIndication></xd:SignedDataObjectProperties></xd:SignedProperties>";

            var hashRaw = HashAndEncodeBytes(Encoding.UTF8.GetBytes(propertiesRaw));
            var hashId = HashAndEncodeBytes(Encoding.UTF8.GetBytes(propertiesFromId));

            Reference signedPropertiesReference = new()
            {
                Type = "http://uri.etsi.org/01903#SignedProperties",
                Uri = "#idSignedProperties"
            };
            XmlDsigC14NTransform c14Transform = new();

            XmlDsigC14NTransform testTransform = new(true);

            var originDoc = new XmlDocument() { PreserveWhitespace = true };
            originDoc.LoadXml(propertiesRaw);

            testTransform.LoadInput(originDoc.GetElementsByTagName("xd:SignedProperties")[0].ChildNodes);

            var output2 = Convert.ToBase64String(testTransform.GetDigestedOutput(SHA1.Create()));

            signedPropertiesReference.AddTransform(c14Transform);
            signedPropertiesReference.DigestMethod = DigestMethods.SHA1;

            var doc = new XmlDocument();

            RSACryptoServiceProvider rsaKey = new();
            RSA key = RSA.Create();

            ExcelSignedXml signedXml = new(doc)
            {
                SigningKey = rsaKey,
            };

            signedXml.Signature.Id = "idPackageSignature";
            signedXml.SignedInfo.CanonicalizationMethod = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            DataObject anObject = new DataObject();

            var element = originDoc.GetElementsByTagName("xd:SignedProperties")[0];

            anObject.LoadXml(originDoc.DocumentElement);
            var test1 = anObject.GetXml();

            signedXml.AddReference(signedPropertiesReference);
            signedXml.AddObject(anObject);

            signedXml.ComputeSignature();

            var testXml = signedXml.GetXml();
        }

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
        public void ReadFileWithSignatureAndSignatureLine()
        {
            using (var pck = OpenTemplatePackage("StampSignature.xlsx"))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets[0];

                var signature = wb.DigitialSignatures[0];
                var vmlDrawings = ws.VmlDrawings;

                foreach (var drawrin in vmlDrawings)
                {
                    var baseDraw = drawrin;
                }

                var drawings = ws.Drawings;

                //var ws = pck.Workbook.Worksheets.Add("ws_SignatureLine");
                //var wb = pck.Workbook;
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

                //var pic = ws.Drawings.ad("Landscape", new FileInfo(@"C:\Users\OssianEdström\Pictures\webp.jpg"));
                //pic.SetPosition(2, 0, 1, 0);

                SaveAndCleanup(pck);
            }
        }

        //[TestMethod]
        //public void SignFileWithSignatureLine()
        //{
        //    using (var pck = OpenPackage("DigSig_SignatureLine.xlsx", true))
        //    {
        //        var ws = pck.Workbook.Worksheets.Add("ws_SignatureLine");
        //        var wb = pck.Workbook;
        //        ws.Drawings.
        //        ExcelVmlDrawingBase
        //    }
        //}


        [TestMethod]
        public void SignFile()
        {
            using (var pck = OpenTemplatePackage("combineddatareport.xlsx"))
            {
                var wb = pck.Workbook;

                wb.FullCalcOnLoad = false;

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var digSig = wb.DigitialSignatures.AddSignature(store.Certificates[1]);
                var info = digSig.SignerInformation;

                info.SignerRoleTitle = "A Title";
                info.Address1 = "Some";
                info.Address2 = "Where";
                info.ZIPorPostalCode = "Over";
                info.City = "The";
                info.CountryOrRegion = "Rainbow";
                info.StateOrProvince = "WayUpHigh";

                //digSig.additionalSignInfo.SignerRoleTitle = "Developer";

                //digSig.additionalSignInfo.productionPlace.Address1 = "ssa";

                ////var prodPlace = digSig.additionalSignInfo.productionPlace;

                ////var prodPlace = new ProductionPlace();

                ////prodPlace.Address1 = "Some";
                ////prodPlace.Address2 = "Where";
                ////prodPlace.ZIPorPostalCode = "Over";
                ////prodPlace.City = "The";
                ////prodPlace.CountryOrRegion = "Rainbow";
                ////prodPlace.StateOrProvince = "WayUpHigh";

                ////digSig.additionalSignInfo.productionPlace = prodPlace;

                SaveAndCleanup(pck);
            }
        }
        [TestMethod]
        public void ReadSignedFileWithAdditionalInfo()
        {
            using (var pck = OpenPackage("combineddatareport.xlsx"))
            {
                var wb = pck.Workbook;
                var signerInformation = wb.DigitialSignatures[0].SignerInformation;
            }
        }

        [TestMethod]
        public void SignFileBig3()
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
        public void HashEncMeta()
        {
            var data = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\"><metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\"  copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\" /></metadataTypes><futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\"><xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"1\"/></ext></extLst></bk></futureMetadata><cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>";
            var byteData = Encoding.UTF8.GetBytes(data);
            var res = HashAndEncodeBytes(byteData);
        }

        [TestMethod]
        public void SignSaveFileWithLOTSOfData()
        {
            using (var pck = OpenTemplatePackage("S610.xlsx"))
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
        public void SHA1Test()
        {
            var bytes = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\sheet1.xml");
            string hash = HashAndEncodeBytes(bytes);

            Assert.AreEqual("5dK/Tn8G0h7N8XnQ6PO7YcqoOWY=", hash);
        }

        [TestMethod]
        public void VerifyTheory()
        {
            string pObject = "<Object Id=\"idOfficeObject\" xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignatureProperties><SignatureProperty Id=\"idOfficeV1Details\" Target=\"#idPackageSignature\"><SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\"><SetupID></SetupID><SignatureText></SignatureText><SignatureImage /><SignatureComments>Forty-two.</SignatureComments><WindowsVersion>10.0</WindowsVersion><OfficeVersion>16.0.17531/26</OfficeVersion><ApplicationVersion>16.0.17531</ApplicationVersion><Monitors>3</Monitors><HorizontalResolution>2560</HorizontalResolution><VerticalResolution>1440</VerticalResolution><ColorDepth>32</ColorDepth><SignatureProviderId>{00000000-0000-0000-0000-000000000000}</SignatureProviderId><SignatureProviderUrl></SignatureProviderUrl><SignatureProviderDetails>9</SignatureProviderDetails><SignatureType>1</SignatureType></SignatureInfoV1></SignatureProperty></SignatureProperties></Object>";
            File.WriteAllText("C:\\epplusTest\\Workbooks\\pObjectTest.xml", pObject, Encoding.UTF8);
            var test = HashAndEncodeBytes(Encoding.UTF8.GetBytes(pObject));

            var readData = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\idOfficeClean.xml");
            var res = HashAndEncodeBytes(readData);

            var officeObj = "<Object Id=\"idOfficeObject\"><SignatureProperties><SignatureProperty Id=\"idOfficeV1Details\" Target=\"#idPackageSignature\"><SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\"><SetupID></SetupID><SignatureText></SignatureText><SignatureImage/><SignatureComments>Forty-two.</SignatureComments><WindowsVersion>10.0</WindowsVersion><OfficeVersion>16.0.17531/26</OfficeVersion><ApplicationVersion>16.0.17531</ApplicationVersion><Monitors>3</Monitors><HorizontalResolution>2560</HorizontalResolution><VerticalResolution>1440</VerticalResolution><ColorDepth>32</ColorDepth><SignatureProviderId>{00000000-0000-0000-0000-000000000000}</SignatureProviderId><SignatureProviderUrl></SignatureProviderUrl><SignatureProviderDetails>9</SignatureProviderDetails><SignatureType>1</SignatureType></SignatureInfoV1></SignatureProperty></SignatureProperties></Object><Object><xd:QualifyingProperties xmlns:xd=\"http://uri.etsi.org/01903/v1.3.2#\" Target=\"#idPackageSignature\"><xd:SignedProperties Id=\"idSignedProperties\"><xd:SignedSignatureProperties><xd:SigningTime>2024-05-27T12:07:02Z</xd:SigningTime><xd:SigningCertificate><xd:Cert><xd:CertDigest><DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/><DigestValue>w9iTMIvTXcdRc9G38Pp1Njb/HPE=</DigestValue></xd:CertDigest><xd:IssuerSerial><X509IssuerName>CN=OssianEdström</X509IssuerName><X509SerialNumber>38225183535545048482234589307877617536</X509SerialNumber></xd:IssuerSerial></xd:Cert></xd:SigningCertificate><xd:SignaturePolicyIdentifier><xd:SignaturePolicyImplied/></xd:SignaturePolicyIdentifier></xd:SignedSignatureProperties><xd:SignedDataObjectProperties><xd:CommitmentTypeIndication><xd:CommitmentTypeId><xd:Identifier>http://uri.etsi.org/01903/v1.2.2#ProofOfOrigin</xd:Identifier><xd:Description>Created and approved this document</xd:Description></xd:CommitmentTypeId><xd:AllSignedDataObjects/><xd:CommitmentTypeQualifiers><xd:CommitmentTypeQualifier>Forty-two.</xd:CommitmentTypeQualifier></xd:CommitmentTypeQualifiers></xd:CommitmentTypeIndication></xd:SignedDataObjectProperties></xd:SignedProperties></xd:QualifyingProperties></Object>";
            byte[] byteTest = Encoding.Default.GetBytes(officeObj);

            var readDataOffice = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\packageObject.xml");
            var officeObject = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\idOfficeObject.xml");

            var testDataOffice = HashAndEncodeBytes(byteTest);
            var testnytt = HashAndEncodeBytes(readDataOffice);
            var officeObjectStuff = HashAndEncodeBytes(officeObject);

            Assert.AreEqual("Dwx/mtIT+lffP980qEOPVRJX41k=", officeObjectStuff);
            Assert.AreEqual("kxA0qm2FwPZvNmtI22ItXRQHlVs=", res);

            byte[] data = Convert.FromBase64String("kxA0qm2FwPZvNmtI22ItXRQHlVs=");

            string decodedString = Encoding.UTF8.GetString(data);
        }

        public string HashAndEncodeBytes(byte[] temp)
        {
            using (var sha1Hash = SHA1.Create())
            {
                var hash = sha1Hash.ComputeHash(temp);
                return Convert.ToBase64String(hash);
            }
        }

        [TestMethod]
        public void DigitallySignDoc()
        {
            using (var p = OpenTemplatePackage("UnsignedWBEmpty.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                foreach (var cert in store.Certificates)
                {
                    var typa = cert.GetType();
                    if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
                    {
                        var sig = p.Workbook.DigitialSignatures.AddSignature(cert);
                        //sig.commitmentType = CommitmentType.CreatedAndApproved;
                        //sig.PurposeForSigning = "I want to";
                        //sig.SignerInformation.
                        //sig.Certificate = cert;
                        break;
                    }
                }

                SaveAndCleanup(p);
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
        public void VBASignTest()
        {
            using (var p = OpenPackage("VbaTest.xlsm", true))
            {
                ExcelPackage pck = new ExcelPackage();
                //Add a worksheet.
                var ws = pck.Workbook.Worksheets.Add("VBA Sample");
                ws.Drawings.AddShape("VBASampleRect", eShapeStyle.RoundRect);
                //Create a vba project             
                pck.Workbook.CreateVBAProject();
                //Now add some code to update the text of the shape...
                var sb = new StringBuilder();
                sb.AppendLine("Private Sub Workbook_Open()");
                sb.AppendLine("    [VBA Sample].Shapes(\"VBASampleRect\").TextEffect.Text = \"This text is set from VBA!\"");
                sb.AppendLine("End Sub");
                pck.Workbook.CodeModule.Code = sb.ToString();

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                foreach (var cert in store.Certificates)
                {
                    if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
                    {
                        pck.Workbook.VbaProject.Signature.Certificate = cert;
                        break;
                    }
                }

                //And Save as xlsm
                pck.SaveAs(new FileInfo(@"C:\epplusTest\Testoutput" + @"\VbaTest.xlsm"));
            }
        }
    }
}