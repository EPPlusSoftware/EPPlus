using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
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
    }
}
