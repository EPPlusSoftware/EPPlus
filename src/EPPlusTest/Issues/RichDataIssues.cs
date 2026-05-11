using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class RichDataIssues : TestBase
    {
        [TestMethod]
        public void PreserveGeoData()
        {
            using var package = OpenTemplatePackage("RichDataPreserve1.xlsx");
            SaveWorkbook("RichDataPreserve1Output.xlsx", package);
        }

        [TestMethod]
        public void PreserveCurrencies()
        {
            using var package = OpenTemplatePackage("RichDataPreserve2.xlsx");
            SaveWorkbook("RichDataPreserve2Output.xlsx", package);
        }

        [TestMethod]
        public void PreserveStocks()
        {
            using var package = OpenTemplatePackage("RichDataPreserve3.xlsx");
            SaveWorkbook("RichDataPreserve3Output.xlsx", package);
        }

        [TestMethod]
        public void RichDataPreserveError1()
        {
            using var package = OpenTemplatePackage("RichDataPreserveError1.xlsx");
            var ws = package.Workbook.Worksheets[0];
            var stockholm = ws.Cells["G1"].Picture.Get();
            var tokyo = ws.Cells["G2"].Picture.Get();
            var ws2 = package.Workbook.Worksheets.Add("Sheet 2");
            ws2.Cells["G1"].Picture.Set(stockholm.GetImageBytes(), "Stockolm");
            ws2.Cells["G2"].Picture.Set(tokyo.GetImageBytes(), "Tokyo");
            SaveWorkbook("RichDataPreserveError1_Output.xlsx", package);
        }

        [TestMethod]
        public void LoadCellsAndCopyLocalImage()
        {
            using var package = OpenTemplatePackage("LocalImageLoadCells.xlsx");
            var ws = package.Workbook.Worksheets[0];
            var stockholm = ws.Cells["D1"].Picture.Get();
            var tokyo = ws.Cells["D2"].Picture.Get();
            var ws2 = package.Workbook.Worksheets.Add("Sheet2");
            ws2.Cells["D1"].Picture.Set(stockholm.GetImageBytes());
            ws2.Cells["D2"].Picture.Set(tokyo.GetImageBytes());
            var stockholm2 = ws2.Cells["D1"].Picture.Get();
            var tokyo2 = ws2.Cells["D2"].Picture.Get();
            SaveWorkbook("LocalImageLoadCells_Output.xlsx", package);
        }

        [TestMethod]
        public void VerifySpillError()
        {
            using (var p = OpenPackage("SpillError.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "RandArray(3,3)";
                ws.Cells["B3"].Value = 4;
                ws.Calculate();
                Assert.IsInstanceOfType(ws.Cells["A1"].Value, typeof(ExcelRichDataErrorValue));
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void VerifyCalcError()
        {
            using (var p = OpenPackage("CalcError.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "Filter(B1:C2,B1:B2<>\"\")";
                ws.Calculate();
                Assert.IsInstanceOfType(ws.Cells["A1"].Value, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)ws.Cells["A1"].Value).Type, eErrorType.Calc);
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void AddingRichTextToFormulasShouldOverwrite()
        {
            using (var pck = OpenPackage("dirtyRT.xlsx"))
            {
                var ws = pck.Workbook.Worksheets.Add("richText");

                ws.Cells["A1"].Value = 1001.1d;
                ws.Cells["C1"].Formula = "ROUND(A1, 1)";
                ws.Cells["B1"].Formula = "\"My favorite number is: \"&TEXT(ROUND(A1,1),\"#,##0.00;(#,##0.00)\")";

                ws.Cells["B1"].RichText.Add("My favorite number is: 1001.1");

                var myFormula = ws.Cells["B1"].Formula;

                Assert.IsTrue(string.IsNullOrEmpty(myFormula));

                ////Set richtext on range that has a formula
                ////This should Throw (or clear Formula)
                //Assert.Throws<InvalidOperationException>(() => ws.Cells["B1"].RichText.Add("My favorite number is: 1001.1"));

                ////This then Results in strange behaviour below when running calculate/IsRichText
                //var cellRange = ws.Cells["B1"];
                //var origRT = cellRange.RichText;

                //ws.Calculate();
                //var afterRt1 = cellRange.Text;
                //var cellRich = ws.Cells["B1"].RichText.Text; //A FRESH reference to the cell, yielding: 1001,10000


                //var OLDCellRich = cellRange.RichText.Text; //Dirty COPY of the cell and its values, yielding: 1001.1\n

                ////This causes the richText of cellRange.RichText to update
                //var myFormula = cellRange.IsRichText;
                ////So does Just LOOKING at the cellRange variable properties In the debugger. It trigger their Getters.
                ////This property changes the actual values of the range when observed in the debugger.
                ////This makes debugging harder and highly confusing both to us and end-users as you may look at a value
                ////See that it is innaccurate and then check it again only to see that it is correct for no discernable reason.

                //var OLDCellRichAfterDebug = cellRange.RichText.Text;
                //Assert.AreEqual(OLDCellRich, OLDCellRichAfterDebug);
            }
        }

        [TestMethod]
        public void AddingRichTextToFormulasShouldThrow()
        {
            using (var pck = OpenPackage("dirtyRT.xlsx"))
            {
                var ws = pck.Workbook.Worksheets.Add("richText");

                ws.Cells["A1"].Value = 1001.1d;
                ws.Cells["C1"].Formula = "ROUND(A1, 1)";
                ws.Cells["B1"].Formula = "\"My favorite number is: \"&TEXT(ROUND(A1,1),\"#,##0.00;(#,##0.00)\")";

                //Set richtext on range that has a formula
                //This should Throw (or clear Formula)
                Assert.Throws<InvalidOperationException>(() => ws.Cells["B1"].RichText.Add("My favorite number is: 1001.1"));

                ////This then Results in strange behaviour below when running calculate/IsRichText
                //var cellRange = ws.Cells["B1"];
                //var origRT = cellRange.RichText;

                //ws.Calculate();
                //var afterRt1 = cellRange.Text;
                //var cellRich = ws.Cells["B1"].RichText.Text; //A FRESH reference to the cell, yielding: 1001,10000


                //var OLDCellRich = cellRange.RichText.Text; //Dirty COPY of the cell and its values, yielding: 1001.1\n

                ////This causes the richText of cellRange.RichText to update
                //var myFormula = cellRange.IsRichText;
                ////So does Just LOOKING at the cellRange variable properties In the debugger. It trigger their Getters.
                ////This property changes the actual values of the range when observed in the debugger.
                ////This makes debugging harder and highly confusing both to us and end-users as you may look at a value
                ////See that it is innaccurate and then check it again only to see that it is correct for no discernable reason.

                //var OLDCellRichAfterDebug = cellRange.RichText.Text;
                //Assert.AreEqual(OLDCellRich, OLDCellRichAfterDebug);
            }
        }

        [TestMethod]
        public void RichTextShouldNotBecomeDirty()
        {
            using (var pck = OpenPackage("dirtyRT.xlsx"))
            {
                var ws = pck.Workbook.Worksheets.Add("richText");

                ws.Cells["A1"].Value = 1001.1d;
                ws.Cells["C1"].Formula = "ROUND(A1, 1)";
                ws.Cells["B1"].Formula = "\"My favorite number is: \"&TEXT(ROUND(A1,1),\"#,##0.00;(#,##0.00)\")";

                //Set richtext on range that has a formula
                //This should clear Formula but does not
                ws.Cells["B1"].RichText.Add("My favorite number is: 1001.1", true);


                //This then Results in strange behaviour below when running calculate/IsRichText
                var cellRange = ws.Cells["B1"];
                var origRT = cellRange.RichText;

                ws.Calculate();
                var afterRt1 = cellRange.Text;
                var cellRich = ws.Cells["B1"].RichText.Text; //A FRESH reference to the cell, yielding: 1001,10000


                var OLDCellRich = cellRange.RichText.Text; //Dirty COPY of the cell and its values, yielding: 1001.1\n

                //This causes the richText of cellRange.RichText to update
                var myFormula = cellRange.IsRichText;
                //So does Just LOOKING at the cellRange variable properties In the debugger. It trigger their Getters.
                //This property changes the actual values of the range when observed in the debugger.
                //This makes debugging harder and highly confusing both to us and end-users as you may look at a value
                //See that it is innaccurate and then check it again only to see that it is correct for no discernable reason.

                var OLDCellRichAfterDebug = cellRange.RichText.Text;

                Assert.AreEqual(OLDCellRich, OLDCellRichAfterDebug);
                Assert.AreNotEqual(cellRich, OLDCellRich);

                var newerRef = ws.Cells["B1"].RichText;

                var refreshedText = "This text should now appear in both pointers";

                cellRange.RichText.Text = refreshedText;

                Assert.AreEqual(refreshedText, newerRef.Text);
                Assert.AreEqual(refreshedText, cellRange.RichText.Text);
                Assert.AreEqual(refreshedText, ws.Cells["B1"].RichText.Text);
            }
        }

    }
}
