using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class CopyDrawingTests : TestBase
    {
        //Sheet 1: 4, 0-3
        //Sheet 2: 9, 0-8
        //Sheet 4: 7, 0-6

        //Copy Shape Tests
        [TestMethod]
        public void CopyShapeSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            Assert.IsTrue(ws0.Drawings.Count < 5);
            ws0.Drawings[0].Copy(ws0, 25, 1);
            Assert.AreEqual(5, ws0._drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyShapeOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws0.Drawings[0].Copy(ws1, 10, 10);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyShapeOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            ws0.Drawings[0].Copy(ws, 10, 10);
            Assert.AreEqual(1, ws.Drawings.Count);
            SaveAndCleanup(p2);
        }
        [TestMethod]
        public void CopyShapeBlipFillTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws0.Drawings[1].Copy(ws1, 10, 20);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }

        //Copy Picture Tests
        [TestMethod]
        public void CopyPictureSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws1.Drawings[0].Copy(ws1, 0, 15);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyPictureOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws0.Drawings.Count < 5);
            ws1.Drawings[0].Copy(ws0, 20, 1);
            Assert.AreEqual(5, ws0.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyPictureOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws0 = p2.Workbook.Worksheets.Add("Sheet1");
            ws1.Drawings[0].Copy(ws0, 1, 1);
            Assert.AreEqual(1, ws0.Drawings.Count);
            SaveAndCleanup(p2);
        }

        //Copy Control Tests
        [TestMethod]
        public void CopyControlSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws1.Drawings[1].Copy(ws1, 25, 20);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyControlOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws2.Drawings.Count < 8);
            ws1.Drawings[1].Copy(ws2, 20, 1);
            Assert.AreEqual(8, ws2.Drawings.Count);
            ws1.Drawings[2].Copy(ws2, 40, 1);
            ws1.Drawings[1].Copy(ws2, 50, 1);
            Assert.AreEqual(10, ws2.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyControlOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            ws1.Drawings[1].Copy(ws, 20, 1);
            ws1.Drawings[2].Copy(ws, 40, 1);
            ws1.Drawings[1].Copy(ws, 50, 1);
            Assert.AreEqual(3, ws.Drawings.Count);
            SaveAndCleanup(p2);
        }

        //Copy Slicer Tests
        [TestMethod]
        public void CopySlicerSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            Assert.IsTrue(ws0.Drawings.Count < 5);
            ws0.Drawings[2].Copy(ws0, 1, 25, 0, 0);
            Assert.AreEqual(5, ws0.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopySlicerOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws2 = p.Workbook.Worksheets[2];
            Assert.IsTrue(ws2.Drawings.Count < 8);
            ws0.Drawings[2].Copy(ws2, 1, 15, 0, 0);
            Assert.AreEqual(8, ws2.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopySlicerOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            var ex = Assert.ThrowsException<InvalidOperationException>(() => ws0.Drawings[2].Copy(ws, 1, 15, 0, 0));
            Assert.AreEqual("Table slicers can't be copied from one workbook to another.", ex.Message);
        }

        //Copy Chart Tests
        [TestMethod]
        public void CopyChartSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            Assert.IsTrue(ws2.Drawings.Count < 8);
            ws2.Drawings[0].Copy(ws2, 20, 1);
            Assert.AreEqual(8, ws2.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyChartOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws2.Drawings[0].Copy(ws1, 20, 20);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyChartOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            ws2.Drawings[0].Copy(ws, 20, 1);
            Assert.AreEqual(1, ws.Drawings.Count);
            SaveAndCleanup(p2);
        }

        //Copy Group Shape Tests
        [TestMethod]
        public void CopyGroupShapeSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            Assert.IsTrue(ws2.Drawings.Count < 8);
            ws2.Drawings[1].Copy(ws2, 5, 20);
            Assert.AreEqual(8, ws2.Drawings.Count);
            ws2.Drawings[2].Copy(ws2, 5, 25);
            ws2.Drawings[4].Copy(ws2, 5, 30);
            ws2.Drawings[5].Copy(ws2, 5, 35);
            ws2.Drawings[6].Copy(ws2, 5, 40);
            Assert.AreEqual(12, ws2.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyGroupShapeOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws2.Drawings[1].Copy(ws1, 5, 20);
            Assert.AreEqual(10, ws1.Drawings.Count);
            ws2.Drawings[2].Copy(ws1, 5, 25);
            ws2.Drawings[4].Copy(ws1, 5, 30);
            ws2.Drawings[5].Copy(ws1, 5, 35);
            ws2.Drawings[6].Copy(ws1, 5, 40);
            Assert.AreEqual(14, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyGroupShapeOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            ws2.Drawings[1].Copy(ws, 1, 1);
            ws2.Drawings[2].Copy(ws, 1, 5);
            ws2.Drawings[4].Copy(ws, 5, 10);
            ws2.Drawings[5].Copy(ws, 5, 15);
            Assert.AreEqual(4, ws.Drawings.Count);
            SaveAndCleanup(p2);
        }
        [TestMethod]
        public void CopySlicerInGroupShapeOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            var ex = Assert.ThrowsException<InvalidOperationException>(() => ws2.Drawings[6].Copy(ws, 5, 40));
            Assert.AreEqual("Table slicers can't be copied from one workbook to another.", ex.Message);
        }

        //Copy PivotTable Slicers
        [TestMethod]
        public void CopyPivotSlicerSameWorksheet()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws3 = p.Workbook.Worksheets[3];
            Assert.IsTrue(ws3.Drawings.Count < 4);
            ws3.Drawings[1].Copy(ws3, 1, 15, 0, 0);
            Assert.AreEqual(4, ws3.Drawings.Count);
            SaveAndCleanup(p);
        }

        //Copy Range
        [TestMethod]
        public void CopyDrawingsRange()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws3 = p.Workbook.Worksheets[2];
            ws3.Cells["A1:Z50"].Copy(ws3.Cells["AA1:AZ50"]);
            SaveAndCleanup(p);
        }


        private class Item
        {
            public string? Name { get; set; }

            public int Value { get; set; }
        }

        [TestMethod]
        public void issue1475()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");


            IEnumerable<Item> _items = new List<Item>() {
                new Item { Name = "Bob", Value = 3 },
                new Item { Name = "Lisa", Value = 8 },
                new Item { Name = "Steve", Value = 5 },
                new Item { Name = "Phil", Value = 2 },
            };

            var range = ws.Cells["A1"].LoadFromCollection(_items, true, TableStyles.Dark1);
            var chart = ws.Drawings.AddLineChart("LineChartWithDroplines", eLineChartType.Line);
            var serie = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0));
            serie.Header = "Order Value";
            chart.SetPosition(0, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart With Droplines";
            chart.AddDropLines();
            chart.DropLine.Border.Width = 2;
            //Set style 12
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle12);


            var cpyWs = package.Workbook.Worksheets.Add("Copy", ws);
            cpyWs.View.TabSelected = false;
            package.SaveAs("C:\\epplusTest\\Testoutput\\i1475.xlsx");
        }
    }
}
