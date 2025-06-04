using FontLab1;
using FontLab1.GenericMeasurements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.PDF.PdfSettings;

namespace EPPlusTest.PDF.FontLabs1
{
    [TestClass]
    public class FontLabs1Tests : TestBase
    {
        [TestMethod]
        public void ReadFontsFromSystem()
        {
            PdfPageSettings pageSettings = new PdfPageSettings();
            //TtfFont arialData = GenericFonts.GetFontData("Arial");
            TtfFont aptosData = GenericFonts.GetFontData(pageSettings, "Aptos Narrow");
        }
    }
}
