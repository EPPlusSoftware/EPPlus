using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.VBA;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;

namespace EPPlusTest.Drawing.Control
{
    [TestClass]
    public class DrawingGroupingTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        static ExcelVBAModule _codeModule;
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
    }
}
