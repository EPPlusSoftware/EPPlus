using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.VBA;
using System;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;

namespace EPPlusTest.Drawing.Grouping
{
    [TestClass]
    public class DrawingGroupingTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DrawingGrouping.xlsx",true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void Group_GroupBoxWithRadioButtonsTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("GroupBox");
            var ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
            ctrl.Text = "Groupbox 1";
            ctrl.SetPosition(480, 80);
            ctrl.SetSize(200, 120);

            _ws.Cells["G1"].Value = "Linked Groupbox";            
            ctrl.LinkedCell = _ws.Cells["G1"];

            var r1 = _ws.Drawings.AddRadioButtonControl("Option Button 1");
            r1.SetPosition(500, 100);
            r1.SetSize(100, 25);
            var r2 = _ws.Drawings.AddRadioButtonControl("Option Button 2");
            r2.SetPosition(530, 100);
            r2.SetSize(100, 25);
            var r3 = _ws.Drawings.AddRadioButtonControl("Option Button 3");
            r3.SetPosition(560, 100);
            r3.SetSize(100, 25);
            r3.FirstButton = true;

            ctrl.Group(r1, r2, r3);

        }
        [TestMethod]
        public void Group_SingleDrawing()
        {
            _ws = _pck.Workbook.Worksheets.Add("SingleDrawing");
            var ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
            ctrl.SetPosition(480, 80);
            ctrl.SetSize(200, 120);

            ctrl.Group();
        }
        [TestMethod]
        public void Group_AddControlViaGroupShape()
        {
            _ws = _pck.Workbook.Worksheets.Add("AddViaGroupShape");
            var ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
            ctrl.SetPosition(480, 80);
            ctrl.SetSize(200, 120);

            var r1 = _ws.Drawings.AddRadioButtonControl("Option Button 1");
            r1.SetPosition(500, 100);
            r1.SetSize(100, 25);

            var group = ctrl.Group();
            group.Drawings.Add(r1);
        }
        [TestMethod]
        public void UnGroup_SingleDrawing()
        {
            _ws = _pck.Workbook.Worksheets.Add("UnGroupSingleDrawing");
            var ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
            ctrl.SetPosition(480, 80);
            ctrl.SetSize(200, 120);
            
            ctrl.Group();
            ctrl.UnGroup();
        }
        [TestMethod]
        public void UnGroup_GroupBoxWithRadioButtonsTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("UnGroupAllDrawings");
            var ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
            ctrl.Text = "Groupbox 1";
            ctrl.SetPosition(480, 80);
            ctrl.SetSize(200, 120);

            _ws.Cells["G1"].Value = "Linked Groupbox";
            ctrl.LinkedCell = _ws.Cells["G1"];

            var r1 = _ws.Drawings.AddRadioButtonControl("Option Button 1");
            r1.SetPosition(500, 100);
            r1.SetSize(100, 25);
            var r2 = _ws.Drawings.AddRadioButtonControl("Option Button 2");
            r2.SetPosition(530, 100);
            r2.SetSize(100, 25);
            var r3 = _ws.Drawings.AddRadioButtonControl("Option Button 3");
            r3.SetPosition(560, 100);
            r3.SetSize(100, 25);
            r3.FirstButton = true;

            var g=ctrl.Group(r1, r2, r3);

            g.SetPosition(100, 100);    //Move whole group

            r1.UnGroup(false);
        }
        [TestMethod]
        public void Group_GroupIntoGroupTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("GroupIntoGroup");
            var ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
            ctrl.Text = "Groupbox 1";
            ctrl.SetPosition(480, 80);
            ctrl.SetSize(200, 120);

            _ws.Cells["G1"].Value = "Linked Groupbox";
            ctrl.LinkedCell = _ws.Cells["G1"];

            var r1 = _ws.Drawings.AddRadioButtonControl("Option Button 1");
            r1.SetPosition(500, 100);
            r1.SetSize(100, 25);
            var g = ctrl.Group(r1);
            var r2 = _ws.Drawings.AddRadioButtonControl("Option Button 2");
            r2.SetPosition(530, 100);
            r2.SetSize(100, 25);
            g = ctrl.Group(r2);
            var r3 = _ws.Drawings.AddRadioButtonControl("Option Button 3");
            r3.SetPosition(560, 100);
            r3.SetSize(100, 25);
            r3.FirstButton = true;
            g.Drawings.Add(r3);
        }
        [TestMethod]
        public void Group_ShapeAndChart()
        {
            _ws = _pck.Workbook.Worksheets.Add("ShapeAndChart");
            var chart = _ws.Drawings.AddLineChart("LineChart 1", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);

            var shape = _ws.Drawings.AddShape("Shape 1", eShapeStyle.Octagon);
            shape.SetPosition(200, 200);

            chart.Group(shape);
        }

        [TestMethod]
        public void Group_PictureAndSlicer()
        {
            _ws = _pck.Workbook.Worksheets.Add("PictureAndSlicer");
            var pic = _ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
            pic.SetPosition(400, 400);


            var tbl = _ws.Tables.Add(_ws.Cells["A1:B2"], "Table1");       
            var slicer = tbl.Columns[0].AddSlicer();
            slicer.SetPosition(200, 200);

            pic.Group(slicer);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Group_GroupIntoOthereWorksheetShouldFailText()
        {
            using (var p = new ExcelPackage())
            {
                var ws1 = p.Workbook.Worksheets.Add("Sheet1");
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");
                var ctrl1 = (ExcelControlGroupBox)ws1.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
                ctrl1.Text = "Groupbox 1";
                ctrl1.SetPosition(480, 80);
                ctrl1.SetSize(200, 120);

                var ctrl2 = (ExcelControlGroupBox)ws1.Drawings.AddControl("GroupBox 2", eControlType.GroupBox);
                ctrl2.Text = "Groupbox 2";
                ctrl2.SetPosition(480, 400);
                ctrl2.SetSize(200, 120);

                var r1 = ws2.Drawings.AddRadioButtonControl("Option Button 1");
                r1.SetPosition(500, 100);
                r1.SetSize(100, 25);
                var g = ctrl1.Group(r1);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Group_GroupIntoOtherGroupShouldFailTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("UnGroupAllDrawings");
                var ctrl1 = (ExcelControlGroupBox)ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
                ctrl1.Text = "Groupbox 1";
                ctrl1.SetPosition(480, 80);
                ctrl1.SetSize(200, 120);

                var ctrl2 = (ExcelControlGroupBox)ws.Drawings.AddControl("GroupBox 2", eControlType.GroupBox);
                ctrl2.Text = "Groupbox 2";
                ctrl2.SetPosition(480, 400);
                ctrl2.SetSize(200, 120);

                var r1 = ws.Drawings.AddRadioButtonControl("Option Button 1");
                r1.SetPosition(500, 100);
                r1.SetSize(100, 25);
                var g = ctrl1.Group(r1);

                ctrl2.Group(r1);
            }
        }
    }
}
