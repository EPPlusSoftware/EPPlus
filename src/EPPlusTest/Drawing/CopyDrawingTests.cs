using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class CopyDrawingTests : TestBase
    {
        [TestMethod]
        public void CopyShapeTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            ws0.Drawings[0].Copy(ws1, 10, 10);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void CopyShapeBlipFillTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            ws0.Drawings[1].Copy(ws1, 10, 20);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void CopyPictureTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            ws1.Drawings[0].Copy(ws0, 20, 1);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void CopyPictureTestExternal()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[1];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws1 = p2.Workbook.Worksheets.Add("Sheet1");
            ws0.Drawings[0].Copy(ws1, 20, 1);
            ws0.Drawings[0].Copy(ws1, 20, 10);
            SaveAndCleanup(p2);
        }

        [TestMethod]
        public void CopyControlTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            ws1.Drawings[1].Copy(ws2, 20, 1);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void CopyChartTestExternal()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var sourceWs = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var targetWs = p2.Workbook.Worksheets.Add("Sheet1");
            sourceWs.Drawings[0].Copy(targetWs, 20, 1);
            SaveAndCleanup(p2);
        }
    }
}
