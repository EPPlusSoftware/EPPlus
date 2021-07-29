using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;

namespace EPPlusTest.Core
{
    [TestClass]
    public class WorksheetColumnTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
         //   _pck = OpenPackage("ColumnTests.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
           // SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void ValidateDefaultWidth()
        {
            using(var p = OpenPackage("columnWidthDefault.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("default");
                var expectedWidth = 9.140625D;
                Assert.AreEqual(expectedWidth, ws.DefaultColWidth);

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateWidthArial()
        {
            foreach(var size in new int[] {6,8,9,10,11,12,14,16,18,20,24,26,28,30,32,36,38,40,42,44,48,72,96,128,256})
            {
                using (var p = OpenPackage($"ColumnWidth\\columnWidthArial{size}.xlsx", true))
                {
                    var ws = p.Workbook.Worksheets.Add($"arial{size}");
                    p.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
                    p.Workbook.Styles.NamedStyles[0].Style.Font.Size = size;

                    //var expectedWidth = 9.140625D;
                    //Assert.AreEqual(expectedWidth, ws.DefaultColWidth);

                    ws.Column(2).Width = ws.DefaultColWidth;
                    SaveAndCleanup(p);
                }
            }
        }
        [TestMethod]
        public void ValidateDefaultWidthArial36()
        {
            using (var p = OpenPackage("columnWidthArial36.xlsx", true))
            {   
                var ws = p.Workbook.Worksheets.Add("arial36");
                p.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
                p.Workbook.Styles.NamedStyles[0].Style.Font.Size = 36;

                //var expectedWidth = 9.140625D;
                //Assert.AreEqual(expectedWidth, ws.DefaultColWidth);

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateDefaultWidthArial72()
        {
            using (var p = OpenPackage("columnWidthArial72.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("arial72");
                p.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
                p.Workbook.Styles.NamedStyles[0].Style.Font.Size = 72;

                //var expectedWidth = 9.140625D;
                //Assert.AreEqual(expectedWidth, ws.DefaultColWidth);

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
            }
        }

    }
}
