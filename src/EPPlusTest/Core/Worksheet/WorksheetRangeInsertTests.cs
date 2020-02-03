using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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
            _pck.Save();
            _pck.Dispose();
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
        public void DeleteColumnEntireDrawing()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DrawingsDeleteEntireColumn");
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

        #endregion
    }
}
