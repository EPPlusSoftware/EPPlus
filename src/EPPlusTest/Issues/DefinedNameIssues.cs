using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class DefinedNameIssues : TestBase
    {
        [TestMethod]
        public void I1238()
        {
            using (var p = OpenTemplatePackage("I1238SlowWorkbook.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.Cells["A1"].Value = 1;
                SaveAndCleanup(p);
            }
        }
		[TestMethod]
		public void s652()
        {
            using (var p = OpenTemplatePackage("s652.xlsm"))
            {
                using var p2 = new ExcelPackage();
                var ws = p.Workbook.Worksheets[0];
                p2.Workbook.Worksheets.Add("New ws", ws);
                SaveWorkbook("s652.xlsx", p2);
            }
		}
        [TestMethod]
        //i1408
        public void VersionName()
        {
            using (var package = OpenTemplatePackage("VersionNameManager.xlsx"))
            {
                var name = package.Workbook.Names.First();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        //i1408
        public void DefinedNamesQuoteError()
        {
            using (var package = OpenPackage("QuoteError.xlsx", true))
            {
                package.Workbook.Worksheets.Add("something");
                package.Workbook.Names.AddValue("Lae_Zel", "zhak vo\"n\"fynh duj");

                var packageTemp = OpenPackage("dummyQuoteWorkbook.xlsx", true);
                packageTemp.Workbook.Worksheets.Add("dummy");
                SaveAndCleanup (packageTemp);

                var file = new FileInfo("C:\\epplusTest\\Testoutput\\dummyQuoteWorkbook.xlsx");

                package.Workbook.ExternalLinks.AddExternalWorkbook(file);

                package.Workbook.Names.AddFormula("编制单位", "\"编制单位：\"&[1]dummyQuoteWorkbook!$D$6");


                //Adding "s\"\"omething" here correctly? results in corrupt worksheet.
                package.Workbook.Names.AddValue("Unended", "s\"omething");


                SaveAndCleanup(package);
            }

            using (var package = OpenPackage("QuoteError.xlsx"))
            {

                Assert.AreEqual("zhak vo\"n\"fynh duj", package.Workbook.Names["Lae_Zel"].Value);
                Assert.AreEqual("\"编制单位：\"&[1]dummyQuoteWorkbook!$D$6", package.Workbook.Names["编制单位"].Formula);
                Assert.AreEqual("s\"omething", package.Workbook.Names["Unended"].Value);

                SaveAndCleanup(package);
            }
        }
    }
}
