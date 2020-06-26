using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Slicer
{
    [TestClass]
    public class SlicerTest : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("SlicerText.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Richtext");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);

            File.Copy(fileName, dirName + "\\SlicerRead.xlsx", true);
        }
        [TestMethod]
        public void ReadSlicer()
        {
            using (var p = OpenTemplatePackage("Slicer.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual(2, ws.Drawings.Count);
                Assert.IsInstanceOfType(ws.Drawings[0], typeof(ExcelSlicer));
                Assert.IsInstanceOfType(ws.Drawings[1], typeof(ExcelSlicer));
                Assert.AreNotEqual("", ws.SlicerRelId);
                Assert.IsNotNull(ws.SlicerXml);

                var slicer = ws.Drawings[0].As.Slicer;
                Assert.AreEqual(eSlicerStyle.None, slicer.Style);
            }
        }
    }
}
