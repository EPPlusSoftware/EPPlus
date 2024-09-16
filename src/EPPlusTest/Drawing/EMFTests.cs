using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing.EMF;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class EMFTests : TestBase
    {
        [TestMethod]
        public void ReadOriginalEmf()
        {
            string path = @"C:\epplusTest\OleTest\EMF\OG_image1.emf";
            EMF emf = new EMF();
            emf.Read(path);
        }

        [TestMethod]
        public void ReadBmpInEmf()
        {
            string path = @"C:\epplusTest\OleTest\EMF\bmp1.emf";
            EMF emf = new EMF();
            emf.Read(path);
        }

        [TestMethod]
        public void ReadEmf()
        {
            //string path = @"C:\epplusTest\OleTest\EMF\pptExample.emf";
            string path = @"C:\epplusTest\OleTest\EMF\OG_image1.emf";
            //string path = @"C:\epplusTest\OleTest\EMF\COOL.emf";
            EMF emf = new EMF();
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
            string path = @"C:\epplusTest\OleTest\EMF\OG_image1.emf";
            EMF emf = new EMF();
            emf.Read(path);
            string imagePath = @"C:\epplusTest\OleTest\EMF\Untitled3.bmp";
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            emf.ChangeImage(imageBytes);
            emf.Save("C:\\epplusTest\\OleTest\\EMF\\ChangeImageTest.emf");
        }

        [TestMethod]
        public void Write100Emf()
        {
            string path = @"C:\epplusTest\OleTest\EMF\image1.emf";
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            EMF[] emfs = new EMF[100];
            for(int i=0; i<100; i++)
            {
                emfs[i] = new EMF();
                emfs[i].Read(path);
                string repeatedAlphabet = string.Concat(Enumerable.Repeat(alphabet, (i / 26) + 1));
                string result = repeatedAlphabet.Substring(0, i);
                emfs[i].UpdateTextRecord(result);
                emfs[i].Save(string.Format( "C:\\epplusTest\\OleTest\\EMF\\EMF2\\Test{0}.emf", i));
            }
        }

        [TestMethod]
        public void Write100Emf2()
        {
            string path = @"C:\epplusTest\OleTest\EMF\image1.emf";
            EMF[] emfs = new EMF[100];
            int i = 0;
            emfs[i] = new EMF();
            emfs[i].Read(path);
            emfs[i].ChangeTextAlignment(TextAlignmentModeFlags.TA_CENTER);
            emfs[i].Save(string.Format("C:\\epplusTest\\OleTest\\EMF\\EMF2\\AlignTest2{0}.emf", i));

            i = 1;
            emfs[i] = new EMF();
            emfs[i].Read(path);
            emfs[i].ChangeTextAlignment(TextAlignmentModeFlags.TA_RIGHT);
            emfs[i].Save(string.Format("C:\\epplusTest\\OleTest\\EMF\\EMF2\\AlignTest2{0}.emf", i));

            i = 2;
            emfs[i] = new EMF();
            emfs[i].Read(path);
            emfs[i].ChangeTextAlignment(TextAlignmentModeFlags.TA_RIGHT | TextAlignmentModeFlags.TA_BOTTOM);
            emfs[i].Save(string.Format("C:\\epplusTest\\OleTest\\EMF\\EMF2\\AlignTest2{0}.emf", i));

            i = 3;
            emfs[i] = new EMF();
            emfs[i].Read(path);
            emfs[i].ChangeTextAlignment(TextAlignmentModeFlags.TA_CENTER | TextAlignmentModeFlags.TA_BOTTOM);
            emfs[i].Save(string.Format("C:\\epplusTest\\OleTest\\EMF\\EMF2\\AlignTest2{0}.emf", i));
        }
    }
}
