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
            using (var p = OpenPackage("firstSheet.xlsx", true))
            {
                for (var i = 0; i < 14; i++)
                {
                    p.Workbook.Worksheets.Add(i.ToString());
                }
                p.Workbook.View.FirstSheet = 5;

                Assert.AreEqual(5, p.Workbook.View.FirstSheet);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateTabRatio()
        {
            using (var p = OpenPackage("tabRatio.xlsx", true))
            {
                for (var i = 0; i < 16; i++)
                {
                    p.Workbook.Worksheets.Add(i.ToString());
                }
                p.Workbook.View.TabRatio = 200;

                Assert.AreEqual(200, p.Workbook.View.TabRatio);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateWorkbookViewVisability()
        {
            using (var p = OpenPackage("visability.xlsx", true))
            {
                for (var i = 0; i < 16; i++)
                {
                    p.Workbook.Worksheets.Add(i.ToString());
                }
                p.Workbook.View.Visibility = eWorkbookVisibility.Visible;
                Assert.AreEqual(eWorkbookVisibility.Visible, p.Workbook.View.Visibility);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateWorkbookViewVisabilityDefault()
        {
            using (var p = OpenPackage("visabilityDefault.xlsx", true))
            {
                for (var i = 0; i < 16; i++)
                {
                    p.Workbook.Worksheets.Add(i.ToString());
                }
                Assert.AreEqual(eWorkbookVisibility.Visible, p.Workbook.View.Visibility);
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ValidateWorkbookViewVisabilityHidden()
        {
            using (var p = OpenPackage("visabilityHidden.xlsx", true))
            {
                for (var i = 0; i < 16; i++)
                {
                    p.Workbook.Worksheets.Add(i.ToString());
                }
                p.Workbook.View.Visibility = eWorkbookVisibility.Hidden;
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ValidateWorkbookViewVisabilityVeryHidden()
        {
            using (var p = OpenPackage("visabilityVeryHidden.xlsx", true))
            {
                for (var i = 0; i < 16; i++)
                {
                    p.Workbook.Worksheets.Add(i.ToString());
                }
                p.Workbook.View.Visibility = eWorkbookVisibility.VeryHidden;
                Assert.AreEqual(eWorkbookVisibility.VeryHidden, p.Workbook.View.Visibility);
                SaveAndCleanup(p);
            }
        }
    }
}
