using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;

namespace EPPlusTest.Core
{
    [TestClass]
    public class WorksheetColumnTest : TestBase
    {
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
        public void ValidateWidthHeeboLight()
        {
            CreateNormalFontsFiles("Heebo Light");
        }
        [TestMethod, Ignore]
        public void ValidateWidthVerdana()
        {
            CreateNormalFontsFiles("Verdana");
        }
        [TestMethod, Ignore]
        public void ValidateWidthArial()
        {
            CreateNormalFontsFiles("Arial");
        }
        [TestMethod, Ignore]
        public void ValidateWidthCalibri()
        {
            CreateNormalFontsFiles("Calibri");
        }
        [TestMethod, Ignore]
        public void ValidateWidthTimesNewRoman()
        {
            CreateNormalFontsFiles("Times New Roman");
        }        
        private static void CreateNormalFontsFiles(string fontName)
        {
            var fontNameNoSpace = fontName.Replace(" ", "");
            foreach (var size in new int[] { 6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 48, 72, 96, 128, 256 })
            {
                using (var p = OpenPackage($"ColumnWidth\\columnWidth{fontNameNoSpace}{size}.xlsx", true))
                {
                    var ws = p.Workbook.Worksheets.Add($"{fontNameNoSpace}{size}");
                    p.Workbook.Styles.NamedStyles[0].Style.Font.Name = fontName;
                    p.Workbook.Styles.NamedStyles[0].Style.Font.Size = size;

                    ws.Column(2).Width = ws.DefaultColWidth;
                    SaveAndCleanup(p);
                }
            }
        }
        [TestMethod]
        public void ValidateAutoFitWidthNormalArial28()
        {
            using (var p = OpenPackage($"columnWidthArial28.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add($"arial28");
                p.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
                p.Workbook.Styles.NamedStyles[0].Style.Font.Size = 28;

                ws.Cells["A1"].Value = "12345678";
                ws.Column(1).AutoFit();

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
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

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
            }
        }

    }
}
