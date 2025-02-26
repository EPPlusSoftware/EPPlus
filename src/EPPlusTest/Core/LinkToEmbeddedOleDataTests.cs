using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.OleObject;
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
            var b1 = ole.GetEmbeddedObjectBytes();
            var pp = ole.GetEmbeddedPackage();
            var b2 = ole.GetEmbeddedObjectBytes();
            ole.SetEmbeddedPackage(pp);
            var b3 = ole.GetEmbeddedObjectBytes();
            ws.Cells["A5"].CreateArrayFormula("[1]!'!Sheet1!Object 1!Sheet1!R2C3'");
            SaveAndCleanup(p);

            var p2 = new ExcelPackage("C:\\epplusTest\\Testoutput\\OleObjectTest_Embed_XLSX.xlsx");
            var ws2 = p2.Workbook.Worksheets[0];
            var ole2 = ws2.Drawings[0] as ExcelOleObject;
            var b4 = ole2.GetEmbeddedObjectBytes();

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

    }
}

/*
 * WHEN EMBEDDING XLSX:
 * 
 * 
 * TODO:
 * write tests.
 * update docs
 * ?
 * 
 * 
 * 
 * External links wiki documentation
 * 
 * ### External Links to Embedded XLSX files.
As of version 8.?, EPPlus has now limited support for links to embedded XLSX files. EPPlus won't use links for calculation but will preservde them and you can write formulas for referencing data in an embedded XLSX file.
There now exsists a method on OLE Objects called var embeddedPackage = GetEmbeddedPackage() that returns the ExcelPackage of the embedded workbook that you can manipulate like any other workbook in EPPlus. To save the changes use SetEmbeddedPackage(embeddedPackage)
to save it.



Formulas have the following format:
[1]!'!Sheet1!Object 1!OleSheet!R7C1'

The [1]! is the index of the externalLink. !Sheet1 is the name of the sheet in the current workbook. !Object 1 is the name of the OLE Object. !OleSheet is the name of the worksheet inside !Object 1. !R7C1 is referecing the 7th row in the 1st Column.
An exmaple use would be ws.Cells["A5"].CreateArrayFormula("[1]!'!Sheet1!Object 1!OleSheet!R7C1'");
 */
