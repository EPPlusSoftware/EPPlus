using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Linq;

namespace EPPlusTest.Style
{
    [TestClass]
    public class ExcelColorTest: TestBase
    {
        [TestMethod]
        public void ExcelIndexColorVerification()
        {
            for(int i = 0; i < ExcelColor.indexedColors.Count(); i++) 
            {
                var colString = "#" + ExcelColor.indexedColorAsColor[i].ToArgb().ToString("x8").ToUpper();
                Assert.AreEqual(ExcelColor.indexedColors[i], colString);
            }
        }
    }
}
