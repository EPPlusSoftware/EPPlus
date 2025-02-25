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
        public void GetEmbeddedWorkbookDataTest()
        {
            /*                     Program Id        Path to current workbook, wb                            !Sheet name in wb
             *                                                                                                      !The object index in sheet
                                                                                                                             !Sheet name in object
                                                                                                                                    !row 1 column 1*/
            var embeddedFormula = "Excel.Sheet.12 | 'C:\\epplusTest\\Workbooks\\LinkToEmbeddedOleData.xlsx'!'!Sheet1!Object 1!Sheet1!R1C1'";
            var ef2 = "[1]!'!Sheet1!Object 1!Sheet1!R2C3'";
            var ef3 = "[1]!'!Sheet1!Object 1!Sheet1!R7C1'";

            using var p = OpenTemplatePackage("OleObjectTest_Embed_XLSX - Copy.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            var b = ole.GetEmbeddedObjectBytes();
            //var ext = p.Workbook.ExternalLinks;

            var p2 = ole.GetEmbeddedPackage();
            p2.Workbook.Worksheets[0].Cells["A7"].Value = "Will it blend?";
            ole.SetEmbeddedPackage(p2);
            ws.Cells["A5"].CreateArrayFormula(ef2);// = ef2;//$"IF('[{1}]Sheet1'!A1 = \"Thirkgjs\", 1,0)";
            ws.Cells["C4"].CreateArrayFormula(ef3);
            //ws.Calculate();
            //var v = ws.Cells["A4"].Value;

            p.SaveAs("C:\\epplusTest\\Testoutput\\OleEmbeddedLinkData.xlsx");

        }

        [TestMethod]
        public void ReadExternalXlsxFileTest()
        {
            using var p = new ExcelPackage();
            var wb = p.Workbook;
            var ws = wb.Worksheets.Add("Sheet 1");

            FileInfo extWb = new FileInfo("C:\\epplusTest\\Workbooks\\OleObjectFiles\\MySheet.xlsx");
            var wb2 = p.Workbook.ExternalLinks.AddExternalWorkbook(extWb);
            ws.Cells["A1"].Formula = $"'[{wb2.Index}]Sheet1'!A1";
            ws.Calculate();
            var v = ws.Cells["A1"].Value;
        }
    }
}

/*
 * WHEN EMBEDDING XLSX:
 * 
 * 
 * 
 */
