using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Range.Insert
{
    [TestClass]
    public class RangeInsertDrawingsTests : TestBase
    {
        public static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("WorksheetRangeInsertDeleteDrawings.xlsx", true);
        }
        [ClassCleanup] 
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        #region Row Tests
        [TestMethod]
        public void InsertRowWithDrawing()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsInsertRow");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell",OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var pic =  ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(0, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(0, 0, 22, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            //Act
            ws.InsertRow(1, 1);
            ws.InsertRow(3, 1);

            //Assert
            Assert.AreEqual(1, shape.From.Row);
            Assert.AreEqual(1, pic.From.Row);
            Assert.AreEqual(0, chart.From.Row);

            Assert.AreEqual(12, shape.To.Row);
            Assert.AreEqual(10, chart.To.Row);
        }
        [TestMethod]
        public void InsertRangeWithDrawingFullShiftDown()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsInsertRangeDownFull");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(0, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(0, 0, 22, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            //Act
            ws.Cells["A1:J1"].Insert(eShiftTypeInsert.Down);
            ws.Cells["A3:J3"].Insert(eShiftTypeInsert.Down);

            //Assert
            Assert.AreEqual(1, shape.From.Row);
            Assert.AreEqual(0, pic.From.Row);
            Assert.AreEqual(0, chart.From.Row);

            Assert.AreEqual(12, shape.To.Row);
            Assert.AreEqual(10, chart.To.Row);
        }
        [TestMethod]
        public void InsertRangeWithDrawingFullShiftRight()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsInsertRangeRightFull");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(0, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(0, 0, 22, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            //Act
            ws.Cells["A1:A10"].Insert(eShiftTypeInsert.Right);
            ws.Cells["C1:C10"].Insert(eShiftTypeInsert.Right);

            //Assert
            Assert.AreEqual(1, shape.From.Column);
            Assert.AreEqual(13, pic.From.Column);
            Assert.AreEqual(22, chart.From.Column);

            Assert.AreEqual(12, shape.To.Column);
            Assert.AreEqual(32, chart.To.Column);
        }
        [TestMethod]
        public void InsertRangeWithDrawingPartialShiftDown()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsInsertRangeDownPart");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(0, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(0, 0, 22, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            //Act
            ws.Cells["A1:I1"].Insert(eShiftTypeInsert.Down);
            ws.Cells["B1:J1"].Insert(eShiftTypeInsert.Down);
            ws.Cells["A3:I3"].Insert(eShiftTypeInsert.Down);
            ws.Cells["B3:J3"].Insert(eShiftTypeInsert.Down);

            //Assert
            Assert.AreEqual(0, shape.From.Row);
            Assert.AreEqual(0, pic.From.Row);
            Assert.AreEqual(0, chart.From.Row);

            Assert.AreEqual(10, shape.To.Row);
            Assert.AreEqual(10, chart.To.Row);
        }
        [TestMethod]
        public void InsertRangeWithDrawingPartialShiftRight()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsInsertRangeRightPart");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(0, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(0, 0, 22, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            //Act
            ws.Cells["A1:A2"].Insert(eShiftTypeInsert.Right);
            ws.Cells["A2:A10"].Insert(eShiftTypeInsert.Right);
            ws.Cells["A3:A9"].Insert(eShiftTypeInsert.Right);
            ws.Cells["B3:J3"].Insert(eShiftTypeInsert.Right);

            //Assert
            Assert.AreEqual(0, shape.From.Column);
            Assert.AreEqual(11, pic.From.Column);
            Assert.AreEqual(22, chart.From.Column);

            Assert.AreEqual(10, shape.To.Column);
            Assert.AreEqual(32, chart.To.Column);
        }
        #endregion
        #region Column Tests
        [TestMethod]
        public void InsertColumnWithDrawing()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsInsertColumn");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(20, 0, 0, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(40, 0, 0, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            //Act
            ws.InsertColumn(1, 1);
            ws.InsertColumn(3, 1);

            //Assert
            Assert.AreEqual(1, shape.From.Column);
            Assert.AreEqual(1, pic.From.Column);
            Assert.AreEqual(0, chart.From.Column);

            Assert.AreEqual(12, shape.To.Column);
            Assert.AreEqual(10, chart.To.Column);
        }
        #endregion
    }
}
