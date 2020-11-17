using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.VBA;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;

namespace EPPlusTest.Drawing.Control
{
    [TestClass]
    public class AddControlTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        static ExcelVBAModule _codeModule;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("FormControl.xlsm",true);
            _pck.Workbook.CreateVBAProject();
            _codeModule = _pck.Workbook.VbaProject.Modules.AddModule("ControlEvents");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void AddButtonTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("Buttons");
            var ctrl = _ws.Drawings.AddControl("Button 1", eControlType.Button).As.Control.Button;
            ctrl.Macro = "Button1_Click";
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 100);
            _ws.Cells["A1"].Value = "Linked Button Caption";
            ctrl.LinkedCell = _ws.Cells["A1"];
            _codeModule.Code += "Sub Button1_Click()\r\n  MsgBox \"Clicked Button!!\"\r\nEnd Sub\r\n";
            //ctrl.Text = "Text";
            ctrl.RichText[0].Fill.Color = Color.Red;
            ctrl.RichText[0].Size=18;
            var rt2 = ctrl.RichText.Add(" Blue");
            rt2.Fill.Color = Color.Blue;
            rt2.Size = 24;

            ctrl.Margin.Automatic = false;
            ctrl.Margin.SetUnit(eMeasurementUnits.Millimeters);
            ctrl.Margin.LeftMargin = 1;
            ctrl.Margin.TopMargin = 2;
            ctrl.Margin.RightMargin = 3;
            ctrl.Margin.BottomMargin = 4;

            ctrl.TextAnchor = eTextAnchoringType.Distributed;
            ctrl.TextAlignment = eTextAlignment.Right;

            ctrl.LayoutFlow = eLayoutFlow.VerticalIdeographic;
            ctrl.Orientation = eShapeOrienation.TopToBottom;
            ctrl.ReadingOrder = eReadingOrder.LeftToRight;
            ctrl.AutomaticSize = true;
            
            Assert.AreEqual(eEditAs.Absolute ,ctrl.EditAs);
            Assert.AreEqual("A1", ctrl.FmlaTxbx.Address);
        }
        [TestMethod]
        public void AddCheckboxTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("Checkbox");
            var ctrl = _ws.Drawings.AddControl("Checkbox 1", eControlType.CheckBox).As.Control.CheckBox;
            ctrl.Macro = "Checkbox_Click";
            ctrl.Fill.Style = eFillStyle.SolidFill;
            ctrl.Fill.SolidFill.Color.SetPresetColor(ePresetColor.Aquamarine);
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 100);
            
            var codeModule = _pck.Workbook.VbaProject.Modules.AddModule("CheckboxCode");
            _codeModule.Code += "Sub Checkbox_Click()\r\n  MsgBox \"Clicked Checkbox!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddRadioButtonTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("RadioButton");
            var ctrl = _ws.Drawings.AddControl("RadioButton 1", eControlType.RadioButton);
            ctrl.Macro = "RadioButton_Click";
            ctrl.SetPosition(500, 100);
            ctrl.SetSize(200, 100);

            var codeModule = _pck.Workbook.VbaProject.Modules.AddModule("RadioButtonCode");
            _codeModule.Code += "Sub RadioButton_Click()\r\n  MsgBox \"Clicked RadioButton!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddDropDownTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("DropDown");
            var ctrl = (ExcelControlDropDown)_ws.Drawings.AddControl("DropDown 1", eControlType.DropDown);
            ctrl.Macro = "DropDown_Click";
            ctrl.SetPosition(500, 100);
            ctrl.SetSize(200, 30);

            _ws.Cells["A1"].Value = 1;
            _ws.Cells["A2"].Value = 2;
            _ws.Cells["A3"].Value = 3;
            _ws.Cells["A4"].Value = 4;

            _ws.Cells["B1"].Value = 3;

            ctrl.InputRange = _ws.Cells["A1:A8"];
            ctrl.LinkedCell = _ws.Cells["B1"];
            ctrl.DropLines = 8;

            _codeModule.Code += "Sub DropDown_Click()\r\n  MsgBox \"Selected DropDown!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddListBoxTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("ListBox");
            var ctrl = (ExcelControlList)_ws.Drawings.AddControl("ListBox 1", eControlType.ListBox);
            ctrl.Macro = "ListBox_Click";
            ctrl.SetPosition(500, 100);
            ctrl.SetSize(200, 100);
            
            _ws.Cells["A1"].Value = 1;
            _ws.Cells["A2"].Value = 2;
            _ws.Cells["A3"].Value = 3;
            _ws.Cells["A4"].Value = 4;

            _ws.Cells["B1"].Value = 3;

            ctrl.InputRange = _ws.Cells["A1:A8"];
            ctrl.LinkedCell = _ws.Cells["B1"];
            
            _codeModule.Code += "Sub ListBox_Click()\r\n  MsgBox \"Selected ListBox!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddLabelTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("Label");
            var ctrl = (ExcelControlLabel)_ws.Drawings.AddControl("Label 1", eControlType.Label);
            ctrl.Macro = "Label_Click";
            ctrl.SetPosition(500, 100);
            ctrl.SetSize(200, 100);

            _ws.Cells["B1"].Value = "Linked Label to B1";

            ctrl.LinkedCell = _ws.Cells["B1"];

            _codeModule.Code += "Sub Label_Click()\r\n  MsgBox \"Selected Label!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddSpinButtonTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("SpinButton");
            var ctrl = (ExcelControlSpinButton)_ws.Drawings.AddControl("SpinButton 1", eControlType.SpinButton);
            ctrl.Macro = "SpinButton_Click";
            ctrl.SetPosition(500, 100);
            ctrl.SetSize(200, 100);

            _ws.Cells["G1"].Value = 3;

            ctrl.LinkedCell = _ws.Cells["G1"];

            _codeModule.Code += "Sub SpinButton_Click()\r\n  MsgBox \"Selected SpinButton!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddGroupBoxTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("GroupBox");
            var ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
            ctrl.Macro = "GroupBox_Click";
            ctrl.SetPosition(500, 100);
            ctrl.SetSize(200, 200);

            _ws.Cells["B1"].Value = "Linked Groupbox";
            
            ctrl.LinkedCell = _ws.Cells["G1"];

            _codeModule.Code += "Sub GroupBox_Click()\r\n  MsgBox \"Clicked GroupBox!!\"\r\nEnd Sub\r\n";
        }
    }
}
