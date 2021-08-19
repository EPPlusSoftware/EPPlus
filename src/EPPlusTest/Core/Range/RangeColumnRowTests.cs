using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Drawing;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeColumnRowTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Range_RowColumn.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void Column_SetWidthBestFitAndStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("ColumnWidth");
            ws.Cells["A1:E5"].EntireColumn.Width = 30;
            ws.Cells["A1:E5"].EntireColumn.Style.Fill.SetBackground(Color.Red);

            ws.Cells["C10:C20"].EntireColumn.BestFit=true;

            Assert.AreEqual(30, ws.Cells["A1"].EntireColumn.Width);
            Assert.AreEqual(30, ws.Cells["C1"].EntireColumn.Width);
            Assert.AreEqual(30, ws.Cells["D1"].EntireColumn.Width);
            Assert.IsFalse(ws.Cells["B1"].EntireColumn.BestFit);
            Assert.IsTrue(ws.Cells["C1"].EntireColumn.BestFit);
            Assert.IsFalse(ws.Cells["D1"].EntireColumn.BestFit);

            Assert.AreEqual("", ws.Cells["E100"].EntireColumn.Style.Fill.BackgroundColor.Rgb);
        }

        [TestMethod]
        public void Column_SetPhonetic()
        {
            var ws = _pck.Workbook.Worksheets.Add("Phonetic");
            ws.Cells["A1:E5"].EntireColumn.Phonetic = true;

            Assert.AreEqual(30, ws.Cells["A1"].EntireColumn.Width);
            Assert.AreEqual(30, ws.Cells["C1"].EntireColumn.Width);
            Assert.AreEqual(30, ws.Cells["D1"].EntireColumn.Width);
            Assert.IsFalse(ws.Cells["B1"].EntireColumn.BestFit);
            Assert.IsTrue(ws.Cells["C1"].EntireColumn.BestFit);
            Assert.IsFalse(ws.Cells["D1"].EntireColumn.BestFit);
        }

    }
}
