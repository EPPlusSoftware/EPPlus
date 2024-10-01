using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.EMF;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class EMFTests : TestBase
    {
        //BMP
        [TestMethod]
        public void ReadOriginalEmf()
        {
            string path = @"C:\epplusTest\OleTest\EMF\squresFromExcel.emf";
            EmfImage emf = new EmfImage();
            emf.Read(path);
        }
        [TestMethod]
        public void ReadBmpInEmf()
        {
            string path = @"C:\epplusTest\OleTest\EMF\ChangeImageTest30.emf";
            EmfImage emf = new EmfImage();
            emf.Read(path);
        }
        [TestMethod]
        public void ReadChangedImage_BMP()
        {
            string pathOG = @"C:\epplusTest\OleTest\EMF\squresFromExcel.emf";
            EmfImage emfOG = new EmfImage();
            emfOG.Read(pathOG);
            string path = "C:\\epplusTest\\OleTest\\EMF\\BEST FIXED FILE.emf";
            EmfImage emf = new EmfImage();
            emf.Read(path);
        }

        [TestMethod]
        public void EMFtest()
        {
            GenericImageHandler handler = new GenericImageHandler();
            string path = "C:\\epplusTest\\OleTest\\EMF\\BEST FIXED FILE.emf";
            byte[] emf = File.ReadAllBytes(path);
            MemoryStream ms = new MemoryStream(emf);
            double width, height, horRes, verRes;
            handler.GetImageBounds(ms, ePictureType.Emf, out width, out height, out horRes, out verRes);
        }

        [TestMethod]
        public void ReadEmf()
        {
            //string path = @"C:\epplusTest\OleTest\EMF\pptExample.emf";
            string path = @"C:\epplusTest\OleTest\EMF\OG_image1.emf";
            //string path = @"C:\epplusTest\OleTest\EMF\COOL.emf";
            EmfImage emf = new EmfImage();
            emf.Read(path);
            //emf.Save(@"C:\epplusTest\OleTest\EMF\newSig1.emf");
            //EMF cool = new EMF();
            //cool.Read(Coolpath);

            //emf.CreateTextRecord("heyo bingus What is it?"); //MÅSTE HAR FILLER TECKEN FÖR SPACING ANNARS BLIR DET KORRUPT!
            //emf.Save(@"C:\epplusTest\OleTest\EMF\image1_COOL8.emf");

        }

        [TestMethod]
        public void ChangeImageTest_BMP()
        {
            string path = @"C:\epplusTest\OleTest\EMF\squresFromExcel.emf";
            EmfImage emf = new EmfImage();
            emf.Read(path);
            string imagePath = @"C:\epplusTest\OleTest\Icons\PDFIcon.bmp";
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            emf.ChangeImage(imageBytes);
            emf.Save("C:\\epplusTest\\OleTest\\EMF\\ChangeImageTest51.emf");
        }

        [TestMethod]
        public void ReadWriteImage_BMP()
        {
            string path = @"C:\epplusTest\OleTest\EMF\bmp1.emf";
            EmfImage emf = new EmfImage();
            emf.Read(path);
            emf.Save("C:\\epplusTest\\OleTest\\EMF\\ChangeImageTest8.emf");
        }
    }
}
