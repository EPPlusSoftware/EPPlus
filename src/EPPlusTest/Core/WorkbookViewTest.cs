using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;

namespace EPPlusTest.Core
{
    [TestClass]
    public class WorkbookViewTest : TestBase
    {

        [TestMethod]
        public void ValidateFirstSheet()
        {
            using(var p = OpenPackage("firstSheet.xlsx", true))
            {
                for (var i= 0; i <14; i++)
                {
                    p.Workbook.Worksheets.Add(i.ToString());
                }
                p.Workbook.View.FirstSheet = 5;

                Assert.AreEqual(5, p.Workbook.View.FirstSheet);
                SaveAndCleanup(p);
            }
        }        
    }
}
