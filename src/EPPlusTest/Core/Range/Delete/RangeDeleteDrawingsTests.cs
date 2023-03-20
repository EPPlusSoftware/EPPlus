using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace EPPlusTest.Core.Range.Delete
{
    [TestClass]
    public class WorksheetRangeInsertDeleteDrawingsTests : TestBase
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
            Assert.AreEqual(5, shape2.To.RowOff / ExcelDrawing.EMU_PER_PIXEL);

            Assert.AreEqual(2, shape3.From.Row);
            Assert.AreEqual(0, shape3.From.RowOff);
            Assert.AreEqual(5, shape3.To.Row);
            Assert.AreEqual(5, shape3.To.RowOff / ExcelDrawing.EMU_PER_PIXEL);
        }
        #endregion
        #region Column Tests
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

            //Act
            ws.DeleteColumn(1, 1);
            ws.DeleteColumn(3, 1);

            //Assert
            Assert.AreEqual(0, shape.From.Column);
            Assert.AreEqual(0, pic.From.Column);
            Assert.AreEqual(1, chart.From.Column);

            Assert.AreEqual(9, shape.To.Column);
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

            var dv = ws.DataValidations.AddIntegerValidation("C1:D5");
            dv.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.equal;
            dv.Formula.Value = 1;
            //Act
            ws.DeleteColumn(3, 10);

            //Assert
            Assert.AreEqual(2, ws.Drawings.Count);

            Assert.AreEqual(0, shape1.From.Column);
            Assert.AreEqual(2, shape1.To.Column);

            Assert.AreEqual(2, shape3.From.Column);
            Assert.AreEqual(5, shape3.To.Column);
            Assert.AreEqual(0, ws.DataValidations.Count);

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
        #region Range
        [TestMethod]
        public void DeleteRangeWithDrawingFullShiftUp()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsInsertRangeDownFull");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape.SetPosition(2, 0, 0, 0);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(2, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(2, 0, 22, 0);
            chart.EditAs = eEditAs.Absolute;

            //Act
            ws.Cells["A1:J1"].Delete(eShiftTypeDelete.Up);
            ws.Cells["A3:J3"].Delete(eShiftTypeDelete.Up);

            //Assert
            Assert.AreEqual(1, shape.From.Row);
            Assert.AreEqual(2, pic.From.Row);
            Assert.AreEqual(2, chart.From.Row);

            Assert.AreEqual(10, shape.To.Row);
            Assert.AreEqual(12, chart.To.Row);
        }
        [TestMethod]
        public void DeleteRangeWithDrawingFullShiftRight()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteRangeLeftFull");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);
            shape.SetPosition(2, 0, 1, 0);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(2, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(2, 0, 22, 0);
            chart.EditAs = eEditAs.Absolute;

            //Act
            ws.Cells["A1:A12"].Delete(eShiftTypeDelete.Left);
            ws.Cells["C1:C12"].Delete(eShiftTypeDelete.Left);

            //Assert
            Assert.AreEqual(0, shape.From.Column);
            Assert.AreEqual(9, pic.From.Column);
            Assert.AreEqual(22, chart.From.Column);

            Assert.AreEqual(9, shape.To.Column);
            //Assert.AreEqual(picToCol, pic.To.Column);
            Assert.AreEqual(32, chart.To.Column);
        }
        [TestMethod]
        public void DeleteRangeWithDrawingPartialShiftUp()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteRangeUpPart");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(0, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(0, 0, 22, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            //Act
            ws.Cells["A1:I1"].Delete(eShiftTypeDelete.Up);
            ws.Cells["B1:J1"].Delete(eShiftTypeDelete.Up);
            ws.Cells["A3:I3"].Delete(eShiftTypeDelete.Up);
            ws.Cells["B3:J3"].Delete(eShiftTypeDelete.Up);

            //Assert
            Assert.AreEqual(0, shape.From.Row);
            Assert.AreEqual(0, pic.From.Row);
            Assert.AreEqual(0, chart.From.Row);

            Assert.AreEqual(10, shape.To.Row);
            Assert.AreEqual(10, chart.To.Row);
        }
        [TestMethod]
        public void DeleteRangeWithDrawingPartialShiftUpOffset()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteRangeUpPartOff");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            shape.SetPosition(5, 5, 11, 5);

            //Act & Assert
            ws.Cells["A1:X1"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual(4, shape.From.Row);
            Assert.AreEqual(5 * ExcelDrawing.EMU_PER_PIXEL, shape.From.RowOff);

            ws.Cells["A5:X5"].Delete(eShiftTypeDelete.Up);
            Assert.AreEqual(4, shape.From.Row);
            Assert.AreEqual(0, shape.From.RowOff);
        }
        [TestMethod]
        public void DeleteRangeWithDrawingPartialShiftLeftOffset()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteRangeLeftPartOff");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            shape.SetPosition(5, 5, 5, 5);

            //Act & Assert
            ws.Cells["A1:A15"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual(4, shape.From.Column);
            Assert.AreEqual(5 * ExcelDrawing.EMU_PER_PIXEL, shape.From.ColumnOff);

            ws.Cells["E1:E15"].Delete(eShiftTypeDelete.Left);
            Assert.AreEqual(4, shape.From.Column);
            Assert.AreEqual(0, shape.From.ColumnOff);
        }
        [TestMethod]
        public void DeleteRangeWithDrawingPartialShiftLeft()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteRangeLeftPart");
            var shape = ws.Drawings.AddShape("Shape1_TwoCell", OfficeOpenXml.Drawing.eShapeStyle.Rect);

            var pic = ws.Drawings.AddPicture("Picture1_OneCell", Properties.Resources.Test1);
            pic.SetPosition(0, 0, 11, 0);

            var chart = ws.Drawings.AddLineChart("Chart1_TwoCellAbsolute", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            chart.SetPosition(0, 0, 22, 0);
            chart.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;

            //Act
            ws.Cells["A1:A2"].Delete(eShiftTypeDelete.Left);
            ws.Cells["A2:A10"].Delete(eShiftTypeDelete.Left);
            ws.Cells["A3:A9"].Delete(eShiftTypeDelete.Left);
            ws.Cells["B3:J3"].Delete(eShiftTypeDelete.Left);

            //Assert
            Assert.AreEqual(0, shape.From.Column);
            Assert.AreEqual(11, pic.From.Column);
            Assert.AreEqual(22, chart.From.Column);

            Assert.AreEqual(10, shape.To.Column);
            Assert.AreEqual(32, chart.To.Column);
        }
        #endregion
    }
}
