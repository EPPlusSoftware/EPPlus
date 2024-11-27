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

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);
            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);
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
            var resourceFile = GetResourceFile("TemplateForStamp.emf");

            var path = resourceFile.FullName;
            templateEmf.Read(path);

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
        public void ReadEMFPlusRecord()
        {
            var emfTemplate = GetTemplateFile("TestTmp.emf");
            var emf = new EmfImage();
            emf.Read(emfTemplate.FullName);

            //Remove plus Records
            emf.records.RemoveAt(10); //Remove End Emf+ comment
            emf.records.RemoveAt(9); //Remove RestoreDC
            emf.records.RemoveAt(6); //Remove duplicate StretchBlt record
            emf.records.RemoveAt(3); //Remove SaveDC
            emf.records.RemoveAt(2); // Remove Emf+ Comment
            emf.records.RemoveAt(1); // Remove Emf+ header Comment

            emf.Save(GetOutputFile("", "AdjustedTmpFile.emf").FullName);
        }

        [TestMethod]
        public void ReadEMFPlusRecordLong()
        {
            var emfTemplate = GetTemplateFile("SignatureImageExtremeLong.emf");
            var emf = new EmfImage();
            emf.Read(emfTemplate.FullName);

            //Remove plus Records
            emf.records.RemoveAt(13); //Remove End Emf+ comment
            emf.records.RemoveAt(12); //Remove RestoreDC
            emf.records.RemoveAt(10); //Remove unknown private comment
            //emf.records.RemoveAt(8); //Removed duplicate smaller record
            emf.records.RemoveAt(7); //Remove unknown private comment
            emf.records.RemoveAt(3); //Remove SaveDC
            emf.records.RemoveAt(2); // Remove Emf+ Comment
            emf.records.RemoveAt(1); // Remove Emf+ header Comment

            var header = (EMR_HEADER)emf.records[0];

            emf.Save(GetOutputFile("", "AdjustedExtremeLong.emf").FullName);
        }

        [TestMethod]
        public void ReadAdjustedEmfPlusRecord()
        {
            var emfTemplate = GetTemplateFile("AdjustedTmpFile.emf");
            var emf = new EmfImage();
            emf.Read(emfTemplate.FullName);

            var imgRecord = (EMR_STRETCHDIBITS)emf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var bytes = File.ReadAllBytes(GetTemplateFile("TempBitmapLarge.bmp").FullName);

            BitmapHandler handler = new BitmapHandler();
            handler.ReadBitmap(bytes);
            imgRecord.ReadBmpAndUpdateImage(bytes);

            var header = (EMR_HEADER)emf.records[0];

            var width = imgRecord.cxDest;
            var height = imgRecord.cyDest;

            header.Bounds.Right = width; //Max line width
            header.Bounds.Bottom = height; //Max stamp height

            header.Frame.Right = Convert.ToInt32(23.26848249027237 * width);
            header.Frame.Bottom = Convert.ToInt32(23.19254658385093 * height);

            //header.Frame.Right = Convert.ToInt32(23.26848249027237 * header.Bounds.Right);
            //header.Frame.Bottom = Convert.ToInt32(23.19254658385093 * header.Bounds.Bottom);

            //imgRecord.ReadBmpAndUpdateImage(bytes);

            //var pxHeight = handler.informationHeader.pixelHeight;
            //var pxWidth = handler.informationHeader.pixelWidth;

            //var MaxWidth = 205;
            //var MaxHeight = 47;

            //double xRatio = (double)MaxWidth / (double)pxWidth;
            //double yRatio = (double)MaxHeight / (double)pxHeight;

            //double ratio = xRatio < yRatio ? xRatio : yRatio;

            //var cxDest = Convert.ToInt32(pxWidth * ratio);
            //var cyDest = Convert.ToInt32(pxHeight * ratio);

            //imgRecord.ReadBmpAndUpdateImage(bytes);

            //var header = (EMR_HEADER)emf.records[0];
            //header.Bounds.Right = cxDest; //Max line width
            //header.Bounds.Bottom = cyDest; //Max stamp height

            //var yDest = Convert.ToInt32(MaxHeight - (pxHeight * ratio));

            ////header.Frame.Right = 5980; 205px
            ////23.26848249027237 * 0.1 mm for each pixel

            ////header.Frame.Bottom = 3734; 161px
            ////23.19254658385093 * 0.1 mm for each pixel

            //header.Frame.Right = Convert.ToInt32(23.26848249027237 * cxDest);
            //header.Frame.Bottom = Convert.ToInt32(23.19254658385093 * yDest);

            emf.Save(GetOutputFile("", "AdjustedTmpChangedImage.emf").FullName);
        }
    }
}
