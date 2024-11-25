using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class ImageFunctionTests : TestBase
    {
        private class TestHttpsService : IHttpsService
        {
            public byte[] Download(string url)
            {
                return Resources.Png2ByteArray;
            }
        }
        [TestMethod]
        public void ImageTest1()
        {
            using var package = new ExcelPackage();
            package.Settings.ImageFunctionService = new TestHttpsService();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sheet.Calculate();
            var pic = sheet.Cells["A1"].Picture.Get();
        }
    }
}
