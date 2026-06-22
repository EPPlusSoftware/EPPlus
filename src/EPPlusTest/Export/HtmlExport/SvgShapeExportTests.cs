using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Export.HtmlExport;

namespace EPPlusTest.Export.HtmlExport
{
    [TestClass]
    public class SvgShapeExportTests : TestBase
    {
        [TestMethod]
        public void ExportBasicShapeWorksheet()
        {
            int[] values = { 5, 10, 15, 20 };
            using (var package = OpenPackage("HtmlBasicSvgShape.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("ShapeWs");
                ws.Drawings.AddShape("SimpleRect", OfficeOpenXml.Drawing.eShapeStyle.Rect);

                var exporter = ws.Cells["A1:A20"].CreateHtmlExporter();

                exporter.Settings.Drawings.Include = ePictureInclude.IncludeInHtmlOnly;
                exporter.Settings.Drawings.DrawTypeInclude = eDrawingInclude.Shapes;

                var htmlPage = exporter.GetSinglePage();

                GetOutputFile("html", "svgRect");
            }
        }
    }
}
