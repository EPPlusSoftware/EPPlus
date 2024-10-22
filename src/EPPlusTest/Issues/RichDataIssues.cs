using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class RichDataIssues : TestBase
    {
        [TestMethod]
        public void PreserveGeoData()
        {
            using var package = OpenTemplatePackage("RichDataPreserve1.xlsx");
            SaveWorkbook("RichDataPreserve1Output.xlsx", package);
        }

        [TestMethod]
        public void PreserveCurrencies()
        {
            using var package = OpenTemplatePackage("RichDataPreserve2.xlsx");
            SaveWorkbook("RichDataPreserve2Output.xlsx", package);
        }

        [TestMethod]
        public void PreserveStocks()
        {
            using var package = OpenTemplatePackage("RichDataPreserve3.xlsx");
            SaveWorkbook("RichDataPreserve3Output.xlsx", package);
        }
    }
}
