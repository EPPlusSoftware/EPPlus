using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class ConditionalFormattingIssues : TestBase
    {
        [TestMethod]
        public void DatabarNegativesAndFormulasTest()
        {
            var package = OpenTemplatePackage("i1244Databars.xlsm");
            Assert.IsNotNull(package.Workbook);

            SaveAndCleanup(package);
        }

        //Contains blanks when address ref.
        [TestMethod]
        public void ContainsBlanksTest()
        {
            using (var p = OpenTemplatePackage("i1254.xlsx"))
            {

                var sheet = p.Workbook.Worksheets[0];

                sheet.Cells["Z1"].Value = 1;

                sheet.Calculate();

                SaveAndCleanup(p);
            }
        }

        /// <summary>
        /// Saves and disposes a package
        /// </summary>
        /// <param name="pck"></param>

        protected static void SaveAndCleanup(ExcelPackage pck, bool disposePackage = true)
        {
            if (pck.Workbook.Worksheets.Count > 0)
            {
                pck.Save();
            }

            if (disposePackage)
            {
                pck.Dispose();
            }
        }

        [TestMethod]
        public void s725()
        {
            using (var p1 = OpenTemplatePackage("s725.xlsx"))
            {
                var sheet = p1.Workbook.Worksheets[6];
                SaveAndCleanup(p1, false);
                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    var sheet2 = p2.Workbook.Worksheets[6];
                    SaveWorkbook("s725-secondsaveorig.xlsx", p2);
                }
            }
        }
    }
}
