using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.EMF;
using System;
using System.Linq;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.Xml;

namespace EPPlusTest
{
    [TestClass]
    public class EmfTest : TestBase
    {
        const string emfOutputFolder = "EmfOutputFolder";

        [TestMethod]
        public void ReadWriteTest()
        {
            using (var package = OpenPackage("ReadEmf.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("EmfSheet");

                var path = GetTemplateFile("UnsignedWithDescriptorsOrigBackup.emf").FullName;

                var emf = new EmfImage();
                emf.Read(path);

                var record = (EMR_EXTTEXTOUTW)emf.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).First();

                var outputPath = GetOutputFile(emfOutputFolder, "Generated.emf").FullName;
                emf.Save(outputPath);
            }
        }

        [TestMethod]
        public void ReadWritePreviouslyGeneratedFile()
        {
            using (var package = OpenPackage("ReadEmf.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("EmfSheet");

                var path = GetTemplateFile("Generated.emf").FullName;

                var emf = new EmfImage();
                emf.Read(path);


                var outputPath = GetOutputFile(emfOutputFolder, "GeneratedTwo.emf").FullName;
                emf.Save(outputPath);
            }
        }

        [TestMethod]
        public void ReadEmfAlt()
        {
            var emfImage = new EmfImage();

            var path = GetTemplateFile("LongName.emf").FullName;

            emfImage.Read(path);

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(2);
            var arr = textRecordArr.ToArray();

            var longName = ((EMR_EXTTEXTOUTW)arr[0]);
            var suggestedSigner = ((EMR_EXTTEXTOUTW)arr[1]);

            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);

            var longIndex = emfImage.records.IndexOf(longName);
            var signerIndex = emfImage.records.IndexOf(suggestedSigner);

            var outputPath = GetOutputFile(emfOutputFolder, "ChangeFontOutput.emf").FullName;

            emfImage.Save(outputPath);
        }

        [TestMethod]
        public void CheckOGImage()
        {
            var emfImage = new EmfImage();
            var path = GetTemplateFile("OG_image1.emf").FullName;

            emfImage.Read(path);

            var textRecord = (EMR_EXTTEXTOUTW)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);
            var fontRecord = (EMR_EXTCREATEFONTINDIRECTW)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);

            Assert.AreEqual("sample.mp3", textRecord.Text);
            Assert.AreEqual("Tahoma", fontRecord.elw.FaceName);
        }

        [TestMethod]
        public void CheckInvalidTemplate()
        {
            var invalidTemplate = new SignatureLineTemplateEmf();
            invalidTemplate.InsertInvalidRecords();
            var records = invalidTemplate.records;

            invalidTemplate.signTextObject.Text = "TemplateSignature";
            invalidTemplate.suggestedSignerObject.Text = "TemplateSigner";
            invalidTemplate.suggestedTitleObject.Text = "TemplateTitle";
            invalidTemplate.SignedBy = "TemplateName";

            var outputPath = GetOutputFile(emfOutputFolder, "InvalidSignatureTemplate2.emf").FullName;

            invalidTemplate.Save(outputPath);
        }

        [TestMethod]
        public void ChangeImageTemplateForStamp()
        {
            var templateEmf = new EmfImage();
            templateEmf.LoadTemplateFromResource("SignatureLineStampTemplate.emf", "OfficeOpenXml.resources.SignatureLineTemplates.zip");

            var templateImage = (EMR_STRETCHDIBITS)templateEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            templateImage.UpdateToImage(GetResourceFile("111By75.bmp").FullName);
            templateEmf.Save(GetOutputFile(emfOutputFolder, "111By75.emf").FullName);

            templateImage.UpdateToImage(GetResourceFile("MaxWidthSignatureStamp.bmp").FullName);
            templateEmf.Save(GetOutputFile(emfOutputFolder, "MaxWidthSignatureStamp.emf").FullName);

            templateImage.UpdateToImage(GetResourceFile("MaxHeightSignatureStamp.bmp").FullName);
            templateEmf.Save(GetOutputFile(emfOutputFolder, "MaxHeightSignatureStamp.emf").FullName);

            templateImage.UpdateToImage(GetResourceFile("5pxSignature.bmp").FullName);
            templateEmf.Save(GetOutputFile(emfOutputFolder, "5pxSignature.emf").FullName);
        }

        [TestMethod]
        public void MaxedOutText()
        {
            var generated = new EmfImage();
            generated.Read(GetTemplateFile("maxedBars.emf").FullName);

            var records = generated.records;

            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Cast<EMR_EXTTEXTOUTW>().ToList();

            textRecords[0].Text = "MMMMMMMMMM||";
            textRecords[0].AdjustReferenceToCenterText(127, 10);

            var outputPath = GetOutputFile("", "MaxTitleGenned.emf").FullName;

            generated.Save(outputPath);
        }

        [TestMethod]
        public void EnsureBitmapCanBeExtractedFromDBitsRecord()
        {
            var emf = new EmfImage();
            emf.Read(GetTemplateFile("5pxSignature.emf").FullName);
            var imgRecord = (EMR_STRETCHDIBITS)emf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var bmpBytes = imgRecord.ExtractedBmp.GetBitMapBytes();

            var outputBmpFile = GetOutputFile("", "5pxGreen.bmp");
            var templateBmpFile = GetTemplateFile("5pxGreen.bmp");

            var templateBytes = File.ReadAllBytes(templateBmpFile.FullName);
            Assert.IsTrue(bmpBytes.SequenceEqual(templateBytes));

            File.WriteAllBytes(outputBmpFile.FullName, bmpBytes);
        }

        [TestMethod]
        public void CheckInValidTemplate()
        {
            var inValidTemplate = new SignatureLineTemplateEmf();
            inValidTemplate.InsertInvalidRecords();
            var records = inValidTemplate.records;

            inValidTemplate.SignText = "IHaveAVeryVeryVeryVerylon";
            inValidTemplate.suggestedSignerObject.Text = "TemplateSigner";
            inValidTemplate.suggestedTitleObject.Text = "TemplateTitle";
            inValidTemplate.SignedBy = "TemplateName";

            var path = GetOutputFile("EmfTests", "InvalidTemplate.emf").FullName;
            inValidTemplate.Save(path);
        }

        [TestMethod]
        public void ReadEmf()
        {
            var emfImage = new EmfImage();
            emfImage.Read(GetTemplateFile("LongName.emf").FullName);

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(2);
            var arr = textRecordArr.ToArray();

            var longName = (EMR_EXTTEXTOUTW)arr[0];
            var suggestedSigner = (EMR_EXTTEXTOUTW)arr[1];

            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);

            var longIndex = emfImage.records.IndexOf(longName);
            var signerIndex = emfImage.records.IndexOf(suggestedSigner);

            emfImage.records[140].data = new byte[] { 3, 0, 0, 0 };

            emfImage.Save(GetOutputFile("EmfTests","ChangedFontTest.emf").FullName);
        }
    }
}
