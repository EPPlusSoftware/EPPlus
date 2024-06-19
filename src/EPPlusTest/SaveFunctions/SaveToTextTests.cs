using FakeItEasy;
using Microsoft.SqlServer.Server;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.SaveFunctions
{
    [TestClass]
    public class SaveToTextTests : TestBase
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
            LoadTestdata(_sheet);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void SaveToTextTextDefault()
        {
            var format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(20,20,25,20);
            format.PaddingCharacterNumeric = '0';
            format.Formats = new string[] { "yyyyMMdd", "","","0.00" };
            _sheet.Cells["A1:D100"].SaveToText(GetOutputFile("TextFiles", "Save1.txt"), format);
        }

        [TestMethod]
        public async Task ToTextFixedWidthAsync()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Value";
            ws.Cells["B1"].Value = 2;
            ws.Cells["C1"].Value = "51%";
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(5, 3, 5);
            var text = await ws.Cells["A1:C1"].ToTextAsync(format);
            Assert.AreEqual("Value  2  51%" + format.EOL, text);
        }

    }
}