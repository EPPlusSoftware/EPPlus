using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class OLETests : TestBase
    {
        [TestMethod]
        public void TestReadEmbeddedObjectBin()
        {
            using var p = OpenTemplatePackage("OLE3.xlsx");
            var ws = p.Workbook.Worksheets[0];

            var ole = ws.Drawings[0] as ExcelOleObject;
        }
    }
}