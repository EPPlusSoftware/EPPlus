using Castle.Core.Resource;
using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using System;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest
{
    [TestClass]
    public class EmfTest : TestBase
    {
        [TestMethod]
        public void ReadWriteTest()
        {
            using (var package = OpenPackage("ReadEmf.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("EmfSheet");

                //var path = "C:\\Users\\OssianEdström\\Downloads\\OG_image1.emf";
                var path = "C:\\epplusTest\\Workbooks\\UnsignedWithDescriptorsOrigBackup.emf";


                var emf = new EmfImage();
                emf.Read(path);

                var record = (EMR_EXTTEXTOUTW)emf.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).First();
                emf.Save("C:\\epplusTest\\Workbooks\\Generated.emf");
            }
        }

        [TestMethod]
        public void ReadWritePreviouslyGeneratedFile()
        {
            using (var package = OpenPackage("ReadEmf.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("EmfSheet");

                var path = "C:\\epplusTest\\Workbooks\\Generated.emf";

                var emf = new EmfImage();
                emf.Read(path);

                emf.Save("C:\\epplusTest\\Workbooks\\GeneratedTwo.emf");
            }
        }

        [TestMethod]
        public void ReadEmfAlt()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\LongName.emf");

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(2);
            var arr = textRecordArr.ToArray();

            var longName = ((EMR_EXTTEXTOUTW)arr[0]);
            var suggestedSigner = ((EMR_EXTTEXTOUTW)arr[1]);

            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);

            var longIndex = emfImage.records.IndexOf(longName);
            var signerIndex = emfImage.records.IndexOf(suggestedSigner);

            //emfImage.records[140].data = new byte[] { 3, 0, 0, 0 };

            emfImage.Save("C:\\epplusTest\\Testoutput\\ChangeFontOutput.emf");
        }

        [TestMethod]
        public void ReadStampExcel()
        {
            var readExcelVersion = new EmfImage();
            readExcelVersion.Read("C:\\epplusTest\\Testoutput\\TemplateBmp.emf");
            readExcelVersion.Save("C:\\epplusTest\\Testoutput\\TemplateResaved.emf");
        }

        [TestMethod]
        public void ReadExtremeWidth()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\RemovedDuplicateWidth.emf");

            var records = emfImage.records;
            var dibits = (EMR_STRETCHDIBITS)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var fileBytes = File.ReadAllBytes(@"C:\Users\OssianEdström\Pictures\LessExtremeLong.bmp");

            dibits.ReadBmpAndUpdateImage(fileBytes);

            dibits.Bounds = new RectLObject(60, 43, 66, 118);

            emfImage.Save("C:\\epplusTest\\Testoutput\\RemovedDuplicateWidthChangedToLong.emf");
        }

        [TestMethod]
        public void ReadExtremeHeight()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\ExtremeHeight.emf");

            var records = emfImage.records;
            var dibits = (EMR_STRETCHDIBITS)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            emfImage.Save("C:\\epplusTest\\Testoutput\\ReSaveExtremeHeight.emf");
        }

        [TestMethod]
        public void ReadStamp2()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\ExtremeWidthHeightRemovedDuplicate.emf");

            var dibits = (EMR_STRETCHDIBITS)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            //dibits.Bounds.Top = 0;
            //dibits.Bounds.Left = 0;

            //dibits.Bounds.Bottom = 78;
            //dibits.Bounds.Right = 128;

            dibits.Bounds = new RectLObject(0, 0, 250, 250);

            var strSrc = @"C:\Users\OssianEdström\Pictures\LessExtremeWide.bmp";
           // var strSrc = @"C:\Users\OssianEdström\Pictures\5PxSignature.bmp"
            var fileBytes = File.ReadAllBytes(strSrc);
            //var handler = new BitmapHandler(fileBytes);

            dibits.ReadBmpAndUpdateImage(fileBytes);

            emfImage.Save("C:\\epplusTest\\Testoutput\\ChangedImageExtremeWide.emf");
        }

        [TestMethod]
        public void ReadStamp()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\ValidStamp.emf");

            var readExcelVersion = new EmfImage();
            readExcelVersion.Read("C:\\epplusTest\\Testoutput\\TemplateBmp.emf");

            var templateRecords = readExcelVersion.records;
            var dibitsTemplate = (EMR_STRETCHDIBITS)readExcelVersion.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var records = emfImage.records;

            var dibits = (EMR_STRETCHDIBITS)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var setWorld = emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var modifyWorld = emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);
            var brushOrgEx = emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_SETBRUSHORGEX);

            var setWorldTemplate = templateRecords.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var modifyWorldTemplate = templateRecords.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);

            setWorld.data = setWorldTemplate.data;
            modifyWorld.data = modifyWorldTemplate.data;

            brushOrgEx.data = new byte[] { 9, 0, 0, 0, 59, 0, 0, 0 };

            var intersectR = new EMR_INTERSECTCLIPRECT();
            intersectR.Clip = new RectLObject(0, 0, 128, 160);
            records.Insert(154, intersectR);

            var fileBytes = File.ReadAllBytes("C:\\Users\\OssianEdström\\Pictures\\ResizedAsExcel.bmp");

            var handler = new BitmapHandler(fileBytes);

            dibits.bitMapHeader = handler.informationHeader;
            dibits.cbBmiSrc = dibits.bitMapHeader.sizeOfHeader;
            dibits.Padding2 = handler.OptionalData;
            dibits.BitsSrc = handler.PixelArray;

            dibits.cxDest = 128;
            dibits.cxSrc = 128;
            dibits.cySrc = 53;
            dibits.cyDest = 53;

            dibits.Bounds = new RectLObject(9, 59, 117, 102);

            emfImage.Save("C:\\epplusTest\\Testoutput\\ValidStampAltered.emf");
        }

        [TestMethod]
        public void CheckOGImage()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\OG_image1.emf");

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);
            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);
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

            validTemplate.Save("C:\\epplusTest\\Testoutput\\ValidSignatureTemplate2.emf");
        }
        [TestMethod]
        public void CheckInvalidTemplate()
        {
            var invalidTemplate = new SignatureLineTemplateInvalid();
            var records = invalidTemplate.records;

            invalidTemplate.signTextObject.Text = "TemplateSignature";
            invalidTemplate.suggestedSignerObject.Text = "TemplateSigner";
            invalidTemplate.suggestedTitleObject.Text = "TemplateTitle";
            invalidTemplate.SignedBy = "TemplateName";

            invalidTemplate.Save("C:\\epplusTest\\Testoutput\\InvalidSignatureTemplate2.emf");
        }

        [TestMethod]
        public void ReadJpg()
        {
            EmfImage jpgEmf = new EmfImage();
            jpgEmf.Read(@"C:\Users\OssianEdström\Pictures\JpgEmf.emf");

            var records = jpgEmf.records;

            var dibitsWidth = (EMR_STRETCHDIBITS)jpgEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
        }

        [TestMethod]
        public void ChangeImageTemplateForStamp()
        {
            var templateEmf = new EmfImage();
            var path = "Resources\\TemplateForStamp.emf";
            templateEmf.Read(path);

            var templateImage = (EMR_STRETCHDIBITS)templateEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            
            templateImage.UpdateToImage("C:\\Users\\OssianEdström\\Pictures\\111By75.bmp");
            templateEmf.Save("C:\\epplusTest\\Testoutput\\ChangedImageEmf.emf");

            templateImage.UpdateToImage("C:\\Users\\OssianEdström\\Documents\\OldPics\\LessExtremeWide.bmp");
            templateEmf.Save("C:\\epplusTest\\Testoutput\\ChangedImageExtremeWide.emf");

            templateImage.UpdateToImage("C:\\Users\\OssianEdström\\Documents\\OldPics\\LessExtremeWide.bmp");
            templateEmf.Save("C:\\epplusTest\\Testoutput\\ChangedImageExtremeWide.emf");

            templateImage.UpdateToImage("C:\\Users\\OssianEdström\\Documents\\OldPics\\LessExtremeLong.bmp");
            templateEmf.Save("C:\\epplusTest\\Testoutput\\ChangedImageExtremeHeight.emf");

            templateImage.UpdateToImage("C:\\Users\\OssianEdström\\Documents\\OldPics\\5pxSignature.bmp");
            templateEmf.Save("C:\\epplusTest\\Testoutput\\5pxSignature.emf");
        }

        [TestMethod]
        public void ReadWorldTransform()
        {
            var emf128 = new EmfImage();
            emf128.Read(@"C:\epplusTest\Testoutput\128EmfFile.emf");

            var dibits128 = (EMR_STRETCHDIBITS)emf128.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var worldTransform = (TransformRecordBase)emf128.records.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var worldTransformModified = (TransformRecordBase)emf128.records.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);

            //worldTransform.xForm.Dy;
            worldTransform.xForm.Dx = 9;
            worldTransform.xForm.M11 = 1;
            worldTransform.xForm.M22 = 1;
            worldTransform.xForm.Dy = 43;

            worldTransformModified.xForm.Dx = 9;
            worldTransformModified.xForm.M11 = 1;
            worldTransformModified.xForm.M22 = 1;
            worldTransformModified.xForm.Dy = 43;

            dibits128.Bounds = new RectLObject(9, 43, 120, 118);

            dibits128.UpdateToImage("C:\\Users\\OssianEdström\\Documents\\Epplus_Repos\\Epplus7\\EPPlus\\src\\EPPlusTest\\Resources\\5PxSignature.bmp");
           // dibits128.ChangeImage2(File.ReadAllBytes("C:\\Users\\OssianEdström\\Pictures\\128Square.bmp"));
            //dibitsExtended.ReadBmpAndUpdateImage(File.ReadAllBytes("C:\\Users\\OssianEdström\\Pictures\\128Square.bmp"));

            emf128.Save(@"C:\epplusTest\Testoutput\TemplateForStamp.emf");
        }


        [TestMethod]
        public void ReadCompareDrawingTest()
        {
            EmfImage wEmf, hEmf, whEmf, Maxed;
            wEmf = new EmfImage();
            hEmf = new EmfImage();
            whEmf = new EmfImage();
            Maxed = new EmfImage();

            wEmf.Read(@"C:\epplusTest\Testoutput\ExtremeWidth.emf");
            hEmf.Read(@"C:\epplusTest\Testoutput\ExtremeHeight.emf");
            whEmf.Read(@"C:\epplusTest\Testoutput\128EmfFile.emf");
            Maxed.Read(@"C:\epplusTest\Testoutput\ValidStampExtended.emf");

            var dibitsWidth = (EMR_STRETCHDIBITS)wEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            var dibitsHeight = (EMR_STRETCHDIBITS)hEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            hEmf.records.Remove(dibitsHeight);
            dibitsHeight = (EMR_STRETCHDIBITS)hEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            var dibitsWidthHeight = (EMR_STRETCHDIBITS)whEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            var dibitsMaxed = (EMR_STRETCHDIBITS)Maxed.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var widthWorld = wEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var widthModify = wEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);

            var heightWorld = hEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var heightModify = hEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);

            var widthHeightWorld = whEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var widthHeightModify = whEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);

            var maxedWorld = Maxed.records.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var maxedModify = Maxed.records.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);

            widthHeightWorld.data = maxedWorld.data;
            widthHeightModify.data = maxedModify.data;

            dibitsWidthHeight.Bounds = new RectLObject(9, 48, 120, 112);
            whEmf.Save(@"C:\epplusTest\Testoutput\whChangedBounds.emf");
            //dibitsHeight.ChangeImage2(File.ReadAllBytes(@"C:\Users\OssianEdström\Pictures\TestBitmap.bmp"));
            //dibitsMaxed.Bounds = new RectLObject(0, 0, 117, 112);
            ////dibitsMaxed.ChangeImage2(File.ReadAllBytes(@"C:\Users\OssianEdström\Pictures\LessExtremeLong.bmp"));

            //hEmf.Save(@"C:\epplusTest\Testoutput\ExtremeHeightChangedWorld.emf");
            //Maxed.Save(@"C:\epplusTest\Testoutput\MaxedResave.emf");
            //var maxedRecords = Maxed.records;
        }
    }
}
