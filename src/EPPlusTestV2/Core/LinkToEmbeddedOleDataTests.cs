using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class LinkToEmbeddedOleDataTests : TestBase
    {
        [TestMethod]
        public void WriteTextToOleObject()
        {
            //Open workbook with external link to embedded object.
            using var p = OpenTemplatePackage("OleObjectTest_Embed_XLSX - Copy.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            string testString = "Will it blend?";

            //Write text to cel in embedded Ole xlsx and save it.
            var p2 = ole.GetEmbeddedPackage();
            p2.Workbook.Worksheets[0].Cells["A7"].Value = testString;
            ole.SetEmbeddedPackage(p2);

            //Test if value was written.
            var p3 = ole.GetEmbeddedPackage();
            Assert.AreEqual(testString, p3.Workbook.Worksheets[0].Cells["A7"].Value);

            p.SaveAs("C:\\epplusTest\\Testoutput\\OleEmbeddedLinkData.xlsx");
        }

        [TestMethod]
        public void WriteLinkToOleFormula_NoExsistingLink()
        {
            //Open workbook without external link to embedded object.
            using var p = OpenTemplatePackage("OleObjectTest_Embed_XLSX.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            ole.CreateLinkToEmbeddedPackage();
            ws.Cells["A5"].CreateArrayFormula("[1]!'!Sheet1!Object 1!Sheet1!R2C3'");
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void WriteLinkToOleFormula_ExsistingLink()
        {
            //Open workbook without external link to embedded object.
            using var p = OpenTemplatePackage("OleObjectTest_Embed_XLSX - Copy.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            ole.GetEmbeddedPackage();
            ws.Cells["A5"].CreateArrayFormula("[1]!'!Sheet1!Object 1!Sheet1!R2C3'");
            SaveAndCleanup(p);
        }


        [TestMethod]
        public void TestInvalidFile()
        {
            //Open package that contains OLE Object that is not an xlsx file.
            using var p = OpenTemplatePackage("OleObjectTest_Embed_DOCX.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;

            Assert.ThrowsException<InvalidOperationException>(() => ole.GetEmbeddedPackage());
        }

    }
}
