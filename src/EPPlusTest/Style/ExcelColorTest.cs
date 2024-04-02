using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace EPPlusTest.Style
{
    [TestClass]
    public class ExcelColorTest: TestBase
    {
        [TestMethod]
        public void ExcelIndexColorVerification()
        {
            using (var package = new ExcelPackage())
            {
                var colors = package.Workbook.Styles.IndexedColors;
                var style = package.Workbook.Styles;

                for (int i = 0; i < colors.Count(); i++)
                {
                    var colString = "#" + style.GetIndexedColor(i).ToArgb().ToString("x8", CultureInfo.InvariantCulture).ToUpper();

                    if (style.GetIndexedColor(i) == Color.Empty)
                    {
                        colString = null;
                    }
                    Assert.AreEqual(style.IndexedColors[i], colString);
                }
            }
        }
    }
}
