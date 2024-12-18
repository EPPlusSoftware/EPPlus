using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    }
}
