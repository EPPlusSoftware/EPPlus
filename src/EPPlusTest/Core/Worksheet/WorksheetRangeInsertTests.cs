using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class WorksheetRangeInsertTests : TestBase
    {
        public static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("WorksheetRangeInsert.xlsx", true);
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

            int picToRow = pic.To.Row;

            //Act
            ws.InsertRow(1, 1);
            ws.InsertRow(3, 1);

            //Assert
            Assert.AreEqual(1, shape.From.Row);
            Assert.AreEqual(1, pic.From.Row);
            Assert.AreEqual(0, chart.From.Row);

            Assert.AreEqual(12, shape.To.Row);
            Assert.AreEqual(picToRow+1, pic.To.Row);
            Assert.AreEqual(10, chart.To.Row);
        }
        [TestMethod]
        public void DeleteRowWithDrawing()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteRow");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape.SetPosition(1, 0, 0, 0);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(1, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(1, 0, 22, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            int picToRow = pic.To.Row;

            //Act
            ws.DeleteRow(1, 1);
            ws.DeleteRow(3, 1);

            //Assert
            Assert.AreEqual(0, shape.From.Row);
            Assert.AreEqual(0, pic.From.Row);
            Assert.AreEqual(1, chart.From.Row);

            Assert.AreEqual(9, shape.To.Row);
            Assert.AreEqual(picToRow - 1, pic.To.Row);
            Assert.AreEqual(11, chart.To.Row);
        }
        [TestMethod]
        public void DeleteRowsEntireDrawing()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteEntireRow");
            var shape1 = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            
            var shape2 = ws.Drawings.AddShape("DeletedShape", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape2.SetPosition(2, 0, 11, 0);

            var shape3 = ws.Drawings.AddShape("Shape3", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape3.SetPosition(5, 0, 22, 0);

            //Act
            ws.DeleteRow(3, 10);  

            //Assert
            Assert.AreEqual(2, ws.Drawings.Count);

            Assert.AreEqual(0, shape1.From.Row);
            Assert.AreEqual(2, shape1.To.Row);

            Assert.AreEqual(2, shape3.From.Row);
            Assert.AreEqual(5, shape3.To.Row);
        }
        [TestMethod]
        public void DeleteRowsDrawingPartialRow()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeletePartialRow");
            var shape1 = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape1.SetPosition(0, 5, 0, 0);

            var shape2 = ws.Drawings.AddShape("PartialShape", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape2.SetPosition(2, 5, 11, 0);

            var shape3 = ws.Drawings.AddShape("Shape3", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape3.SetPosition(5, 5, 22, 0);

            //Act
            ws.DeleteRow(3, 10);

            //Assert
            Assert.AreEqual(3, ws.Drawings.Count);

            Assert.AreEqual(0, shape1.From.Row);
            Assert.AreEqual(5, shape1.From.RowOff / ExcelDrawing.EMU_PER_PIXEL);
            Assert.AreEqual(2, shape1.To.Row);
            Assert.AreEqual(0, shape1.To.RowOff);

            Assert.AreEqual(2, shape2.From.Row);
            Assert.AreEqual(0, shape2.From.RowOff);
            Assert.AreEqual(2, shape2.To.Row);
            Assert.AreEqual(5, shape2.To.RowOff/ExcelDrawing.EMU_PER_PIXEL);

            Assert.AreEqual(2, shape3.From.Row);
            Assert.AreEqual(0, shape3.From.RowOff);
            Assert.AreEqual(5, shape3.To.Row);
            Assert.AreEqual(5, shape3.To.RowOff / ExcelDrawing.EMU_PER_PIXEL);
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

            int picToCol = pic.To.Column;

            //Act
            ws.InsertColumn(1, 1);
            ws.InsertColumn(3, 1);

            //Assert
            Assert.AreEqual(1, shape.From.Column);
            Assert.AreEqual(1, pic.From.Column);
            Assert.AreEqual(0, chart.From.Column);

            Assert.AreEqual(12, shape.To.Column);
            Assert.AreEqual(picToCol + 1, pic.To.Column);
            Assert.AreEqual(10, chart.To.Column);
        }
        [TestMethod]
        public void DeleteColumnWithDrawing()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteColumns");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", eShapeStyle.Rect);
            shape.SetPosition(0, 0, 1, 0);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(11, 0, 1, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(22, 0, 1, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            int picToColumn = pic.To.Column;

            //Act
            ws.DeleteColumn(1, 1);
            ws.DeleteColumn(3, 1);

            //Assert
            Assert.AreEqual(0, shape.From.Column);
            Assert.AreEqual(0, pic.From.Column);
            Assert.AreEqual(1, chart.From.Column);

            Assert.AreEqual(9, shape.To.Column);
            Assert.AreEqual(picToColumn - 1, pic.To.Column);
            Assert.AreEqual(11, chart.To.Column);
        }
        [TestMethod]
        public void DeleteColumnEntireDrawing()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteEntireColumn");
            var shape1 = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var shape2 = ws.Drawings.AddShape("DeletedShape", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape2.SetPosition(11, 0, 2, 0);

            var shape3 = ws.Drawings.AddShape("Shape3", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape3.SetPosition(22, 0, 5, 0);

            //Act
            ws.DeleteColumn(3, 10);

            //Assert
            Assert.AreEqual(2, ws.Drawings.Count);

            Assert.AreEqual(0, shape1.From.Column);
            Assert.AreEqual(2, shape1.To.Column);

            Assert.AreEqual(2, shape3.From.Column);
            Assert.AreEqual(5, shape3.To.Column);
        }
        [TestMethod]
        public void DeleteColumnDrawingPartialColumn()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeletePartialColumn");
            var shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape1.SetPosition(0, 0, 0, 5);

            var shape2 = ws.Drawings.AddShape("PartialShape", eShapeStyle.Rect);
            shape2.SetPosition(11, 0, 2, 5);

            var shape3 = ws.Drawings.AddShape("Shape3", eShapeStyle.Rect);
            shape3.SetPosition(22, 0, 5, 5);

            //Act
            ws.DeleteColumn(3, 10);

            //Assert
            Assert.AreEqual(3, ws.Drawings.Count);

            Assert.AreEqual(0, shape1.From.Column);
            Assert.AreEqual(5, shape1.From.ColumnOff / ExcelDrawing.EMU_PER_PIXEL);
            Assert.AreEqual(2, shape1.To.Column);
            Assert.AreEqual(0, shape1.To.ColumnOff);

            Assert.AreEqual(2, shape2.From.Column);
            Assert.AreEqual(0, shape2.From.ColumnOff);
            Assert.AreEqual(2, shape2.To.Column);
            Assert.AreEqual(5, shape2.To.ColumnOff / ExcelDrawing.EMU_PER_PIXEL);

            Assert.AreEqual(2, shape3.From.Column);
            Assert.AreEqual(0, shape3.From.ColumnOff);
            Assert.AreEqual(5, shape3.To.Column);
            Assert.AreEqual(5, shape3.To.ColumnOff / ExcelDrawing.EMU_PER_PIXEL);
        }

        #endregion
    }
}
