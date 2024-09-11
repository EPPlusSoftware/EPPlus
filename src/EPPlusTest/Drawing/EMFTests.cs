using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.EMF;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class EMFTests : TestBase
    {
        [TestMethod]
        public void ReadEmf()
        {
            //string path = @"C:\epplusTest\OleTest\EMF\pptExample.emf";
            string path = @"C:\epplusTest\OleTest\EMF\image1.emf";
            string Coolpath = @"C:\epplusTest\OleTest\EMF\image1_COOL.emf";
            //string path = @"C:\epplusTest\OleTest\EMF\COOL.emf";
            EMF emf = new EMF();
            emf.Read(path);
            //EMF cool = new EMF();
            //cool.Read(Coolpath);

            emf.CreateTextRecord("COOL TEXT2");
            emf.Save(@"C:\epplusTest\OleTest\EMF\image1_COOL2.emf");

        }

        [TestMethod]
        public void WriteEmf()
        {
            EMF eMF = new EMF();
            eMF.CreateTextRecord("COOL TEXT");
            eMF.Save(@"C:\epplusTest\OleTest\EMF\COOL.emf");
        }
    }
}
