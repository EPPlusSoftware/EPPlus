using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
