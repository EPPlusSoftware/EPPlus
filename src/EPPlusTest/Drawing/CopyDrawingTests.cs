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
        //Copy Shape Tests
        [TestMethod]
        public void CopyShapeSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            ws0.Drawings[0].Copy(ws0, 25, 1);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyShapeOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            ws0.Drawings[0].Copy(ws1, 10, 10);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyShapeOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws1 = p2.Workbook.Worksheets.Add("Sheet1");
            ws0.Drawings[0].Copy(ws1, 10, 10);
            SaveAndCleanup(p2);
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

        //Copy Picture Tests
        [TestMethod]
        public void CopyPictureSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            ws1.Drawings[0].Copy(ws1, 0, 15);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyPictureOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            ws1.Drawings[0].Copy(ws0, 20, 1);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyPictureOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[1];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws1 = p2.Workbook.Worksheets.Add("Sheet1");
            ws0.Drawings[0].Copy(ws1, 1, 1);
            SaveAndCleanup(p2);
        }

        //Copy Control Tests
        [TestMethod]
        public void CopyControlSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            ws1.Drawings[1].Copy(ws1, 25, 20);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyControlOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            ws1.Drawings[1].Copy(ws2, 20, 1);
            ws1.Drawings[2].Copy(ws2, 40, 1);
            ws1.Drawings[1].Copy(ws2, 50, 1);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyControlOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws2 = p2.Workbook.Worksheets.Add("Sheet1");
            ws1.Drawings[1].Copy(ws2, 20, 1);
            ws1.Drawings[2].Copy(ws2, 40, 1);
            ws1.Drawings[1].Copy(ws2, 50, 1);
            SaveAndCleanup(p2);
        }

        //Copy Slicer Tests
        [TestMethod]
        public void CopySlicerSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[0];
            ws1.Drawings[2].Copy(ws1, 1, 25, 0, 0);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopySlicerOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[0];
            var ws3 = p.Workbook.Worksheets[2];
            ws1.Drawings[2].Copy(ws3, 1, 15, 0, 0);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopySlicerOtherWorkbookTest() //Fungerar! Ska slänga ett exception.
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[0];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws2 = p2.Workbook.Worksheets.Add("Sheet1");
            ws1.Drawings[2].Copy(ws2, 1, 15, 0, 0);
            SaveAndCleanup(p2);
        }

        //Copy Chart Tests
        [TestMethod]
        public void CopyChartSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var sourceWs = p.Workbook.Worksheets[2];
            sourceWs.Drawings[0].Copy(sourceWs, 20, 1);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyChartOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var sourceWs = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            sourceWs.Drawings[0].Copy(ws1, 20, 20);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyChartOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var sourceWs = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var targetWs = p2.Workbook.Worksheets.Add("Sheet1");
            sourceWs.Drawings[0].Copy(targetWs, 20, 1);
            SaveAndCleanup(p2);
        }

        //Copy Group Shape Tests
        [TestMethod]
        public void CopyGroupShapeSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws = p.Workbook.Worksheets[2];
            ws.Drawings[1].Copy(ws, 5, 20);
            ws.Drawings[2].Copy(ws, 5, 25);
            ws.Drawings[4].Copy(ws, 5, 30);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyGroupShapeOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            ws.Drawings[1].Copy(ws1, 5, 20);
            ws.Drawings[2].Copy(ws1, 5, 25);
            ws.Drawings[4].Copy(ws1, 5, 30);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyGroupShapeOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var targetWs = p2.Workbook.Worksheets.Add("Sheet1");
            ws.Drawings[1].Copy(targetWs, 1, 1);
            ws.Drawings[2].Copy(targetWs, 1, 5);
            ws.Drawings[4].Copy(targetWs, 5, 30);
            SaveAndCleanup(p2);
        }
    }
}
