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
            using var p = OpenTemplatePackage("OleObjectTest_Embed_DOCX.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;

            Assert.ThrowsException<InvalidOperationException>(() => ole.GetEmbeddedPackage());
        }

    }
}

/*

*Should we link when reading to make user not have to call CreateLinkToembeddedObject()? 
*Copy method update?
*extra image is created...
*xlsm
*xlst


### Linking to Embedded XLSX files.
See the page for OLE Objects.


//Ole Object wiki text

### Link to cell inside an embedded xlsx Ole Object.

While EPPlus won't use the formula for calculation, you can still add formulas that reference data inside an OLE Object. The object must be a valid xlsx file else it might create a corrupt workbook.
To add a formula that references a cell in the embedded xlsx file you must first prepare the file by adding a link to it. You can do it by calling CreateLinkToEmbeddedObject() method on the Ole Object. You can then add
formulas using the following format: [1]!'!Sheet1!Object 1!OleSheet!R7C1'
The [1]! is the index of the externalLink. !Sheet1 is the name of the sheet in the current workbook. !Object 1 is the name of the OLE Object. !OleSheet is the name of the worksheet inside !Object 1. !R7C1 is referecing the 7th row in the 1st Column.
An exmaple use would be ws.Cells["A5"].CreateArrayFormula("[1]!'!Sheet1!Object 1!OleSheet!R7C1'");


### Open an embedded xlsx document for editing.
EPPlus supports opening and editing embedded workbooks. You can use the var embeddedPackage = GetEmbeddedPackage() to get the embedded excel package that you can manipulate like any other workbook in EPPlus.
To save the changes use SetEmbeddedPackage(embeddedPackage) method.

 */
