using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing.EMF;
using System.Collections.Generic;
using System.Linq;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class EMFTests : TestBase
    {
        [TestMethod]
        public void ReadEmf()
        {
            //string path = @"C:\epplusTest\OleTest\EMF\pptExample.emf";
            string path = @"C:\epplusTest\OleTest\EMF\signature1.emf";
            string Coolpath = @"C:\epplusTest\OleTest\EMF\image1_COOL7.emf";
            //string path = @"C:\epplusTest\OleTest\EMF\COOL.emf";
            EMF emf = new EMF();
            emf.Read(path);
            emf.Save(@"C:\epplusTest\OleTest\EMF\newSig1.emf");
            //EMF cool = new EMF();
            //cool.Read(Coolpath);

            //emf.CreateTextRecord("heyo bingus What is it?"); //MÅSTE HAR FILLER TECKEN FÖR SPACING ANNARS BLIR DET KORRUPT!
            //emf.Save(@"C:\epplusTest\OleTest\EMF\image1_COOL8.emf");

        }

        [TestMethod]
        public void WriteEmf()
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
                emfs[i].CreateTextRecord(result, i, -3);
                emfs[i].Save(string.Format( "C:\\epplusTest\\OleTest\\EMF\\EMF2\\Test{0}.emf", i));
            }
        }
    }
}
