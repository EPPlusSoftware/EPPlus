
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.Threading;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeTextTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Range_Text.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("ToTextData");
            var noItems = 100;
            LoadTestdata(_ws, noItems);
            SetDateValues(_ws, noItems);
        }
        [TestMethod]
        public void TextFormatParentheses()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = -100;
                ws.Cells["A1"].Style.Numberformat.Format = "(#,##0)";

                Assert.AreEqual("(100)", ws.Cells["A1"].Text);
                ws.Cells["A1"].Style.Numberformat.Format = "#,##0;(#,##0)";
                Assert.AreEqual("(100)", ws.Cells["A1"].Text);
                ws.Cells["A1"].Style.Numberformat.Format = "#,##0;(#,##0);-";
                Assert.AreEqual("(100)", ws.Cells["A1"].Text);
            }
        }
        [TestMethod]
        public void TextFormatStringWithPercent()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 88;
                ws.Cells["A1"].Style.Numberformat.Format = "0\"%\"";

                Assert.AreEqual("88%", ws.Cells["A1"].Text);
            }
        }
        [TestMethod]
        public void TextFormatWithBlankFormattingNumber()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 88;
                ws.Cells["A2"].Value = 0;
                ws.Cells["A3"].Value = -88;
                ws.Cells["A4"].Value = "String";
                ws.Cells["A1:A4"].Style.Numberformat.Format = ";;";

                Assert.IsNull(ws.Cells["A1"].Text);
                Assert.IsNull(ws.Cells["A2"].Text);
                Assert.IsNull(ws.Cells["A3"].Text);
                Assert.AreEqual("String", ws.Cells["A4"].Text);
            }
        }
        [TestMethod]
        public void TextFormatWithBlankFormattingWithString()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 88;
                ws.Cells["A2"].Value = 0;
                ws.Cells["A3"].Value = -88;
                ws.Cells["A4"].Value = "String";
                ws.Cells["A1:A4"].Style.Numberformat.Format = ";;;";

                Assert.IsNull(ws.Cells["A1"].Text);
                Assert.IsNull(ws.Cells["A2"].Text);
                Assert.IsNull(ws.Cells["A3"].Text);
                Assert.IsNull(ws.Cells["A4"].Text);
            }
        }
        [TestMethod]
        public void NumberFormatWithLanguageCode()
        {
            var dt1 = new DateTime(2021, 3, 30);

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("test");
            ws.Cells[1, 1].Style.Numberformat.Format = "[$-en]d mmmm yyyy";
            ws.Cells[1, 2].Style.Numberformat.Format = "[$-nl]d mmmm yyyy";
            ws.Cells[1, 1, 1, 2].Value = dt1;

            Assert.AreEqual("30 March 2021", ws.Cells[1, 1].Text);
            Assert.AreEqual("30 maart 2021", ws.Cells[1, 2].Text);

            ws.Cells.AutoFitColumns();
        }
        [TestMethod]
        public void ValidateNumberFormatDiffExcelVsNet()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var prevCi = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
                ws.SetValue(1, 1, -0.1);
                ws.SetValue(2, 1, 0);
                ws.SetValue(3, 1, 0.1);
                ws.Cells["A1:A3"].Style.Numberformat.Format = "#,##0;-#,##0;-";
                Assert.AreEqual("-0", ws.Cells["A1"].Text);
                Assert.AreEqual("-", ws.Cells["A2"].Text);
                Assert.AreEqual("0", ws.Cells["A3"].Text);

                ws.Cells["A1:A3"].Style.Numberformat.Format = "#,##0.0;-#,##0.0;-";
                Assert.AreEqual("-0.1", ws.Cells["A1"].Text);
                Assert.AreEqual("-", ws.Cells["A2"].Text);
                Assert.AreEqual("0.1", ws.Cells["A3"].Text);
                Thread.CurrentThread.CurrentCulture = prevCi;
            }
        }
        [TestMethod]
        public void Text()
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = new DateTime(2018, 2, 3);
                ws.Cells["A1"].Style.Numberformat.Format = "d";
                Assert.AreEqual("3", ws.Cells["A1"].Text);
                ws.Cells["A1"].Style.Numberformat.Format = "D";
                Assert.AreEqual("3", ws.Cells["A1"].Text);
                ws.Cells["A1"].Style.Numberformat.Format = "M";
                Assert.AreEqual("2", ws.Cells["A1"].Text);
                ws.Cells["A1"].Style.Numberformat.Format = "Y";
                Assert.AreEqual("18", ws.Cells["A1"].Text);
                ws.Cells["A1"].Style.Numberformat.Format = "YY";
                Assert.AreEqual("18", ws.Cells["A1"].Text);
                ws.Cells["A1"].Style.Numberformat.Format = "YYY";
                Assert.AreEqual("2018", ws.Cells["A1"].Text);
            }
        }
        [TestMethod]
        public void ValudateDateTextWithAMPM()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("dateText");
                ws.Cells["A1"].Value = new DateTime(2021, 7, 6, 9, 29, 0);
                ws.Cells["A1"].Style.Numberformat.Format = "[$-0409]M/d/yyyy h:mm AM/PM";

                Assert.AreEqual("7/6/2021 9:29 AM", ws.Cells["A1"].Text);
            }
        }
        [TestMethod]
        public void ValidateAccountingFormatKr()
        {
            var prevCi = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            var fmt = "_-* #,##0\\ \"kr\"_-;\\-* #,##0\\ \"kr\"_-;_-* \"-\"\\ \"kr\"_-;_-@_-";
            //var fmt2 = "_-* #,##0.00\\ \"kr\"_-;\\-* #,##0.00\\ \"kr\"_-;_-* \"-\"??\\ \"kr\"_-;_-@_-\"";

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("dateText");
                ws.Cells["A1"].Value = 5555;
                ws.Cells["A2"].Value = 0;
                ws.Cells["A3"].Value = -5555;
                ws.Cells["A4"].Value = "Text";
                ws.Cells["A1:A4"].Style.Numberformat.Format = fmt;

                Assert.AreEqual("5,555 kr", ws.Cells["A1"].Text);
                Assert.AreEqual("- kr", ws.Cells["A2"].Text);
                Assert.AreEqual("-5,555 kr", ws.Cells["A3"].Text);
                Assert.AreEqual("Text", ws.Cells["A4"].Text);
            }
            Thread.CurrentThread.CurrentCulture = prevCi;
        }
        [TestMethod]
        public void ValidateAccountingFormatKrWithDecimals()
        {
            var prevCi = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            var fmt = "_-* #,##0.00\\ \"kr\"_-;\\-* #,##0.00\\ \"kr\"_-;_-* \"-\"??\\ \"kr\"_-;_-@_-\"";

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("dateText");
                ws.Cells["A1"].Value = 5555;
                ws.Cells["A2"].Value = 0;
                ws.Cells["A3"].Value = -5555;
                ws.Cells["A4"].Value = "Text";
                ws.Cells["A1:A4"].Style.Numberformat.Format = fmt;

                Assert.AreEqual("5,555.00 kr", ws.Cells["A1"].Text);
                Assert.AreEqual("-   kr", ws.Cells["A2"].Text);
                Assert.AreEqual("-5,555.00 kr", ws.Cells["A3"].Text);
                Assert.AreEqual("Text", ws.Cells["A4"].Text);
            }
            Thread.CurrentThread.CurrentCulture = prevCi;
        }
        [TestMethod]
        public void ValidateDateFormatWithNullAndText()
        {
            var prevCi = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            var fmt = "yyyy/mm/dd;;\"NULL DATE\";\"Invalid date \"@\"";

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("dateText");
                ws.Cells["A1"].Value = new DateTime(2021,2,3);
                ws.Cells["A2"].Value = 0;
                ws.Cells["A3"].Value = -2;
                ws.Cells["A4"].Value = "3/2";
                ws.Cells["A1:A4"].Style.Numberformat.Format = fmt;

                Assert.AreEqual("2021/02/03", ws.Cells["A1"].Text);
                Assert.AreEqual("NULL DATE", ws.Cells["A2"].Text);
                Assert.IsNull(ws.Cells["A3"].Text);
                Assert.AreEqual("Invalid date 3/2", ws.Cells["A4"].Text);
            }
            Thread.CurrentThread.CurrentCulture = prevCi;            
        }

    }
}
