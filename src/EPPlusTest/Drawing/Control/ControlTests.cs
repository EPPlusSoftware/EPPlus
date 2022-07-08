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
    public class ControlTests : TestBase
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
            var ctrl = _ws.Drawings.AddButtonControl("Button 1");
            ctrl.Macro = "Button1_Click";
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 100);
            _ws.Cells["A1"].Value = "Linked Button Caption";
            ctrl.LinkedCell = _ws.Cells["A1"];
            _codeModule.Code += "Sub Button1_Click()\r\n  MsgBox \"Clicked Button!!\"\r\nEnd Sub\r\n";

            ctrl.RichText[0].Fill.Color = Color.Red;
            ctrl.RichText[0].Size = 18;
            var rt2 = ctrl.RichText.Add(" Blue");
            rt2.Fill.Color = Color.Blue;
            rt2.Size = 24;

            ctrl.Margin.Automatic = false;
            ctrl.Margin.SetUnit(eMeasurementUnits.Millimeters);
            ctrl.Margin.LeftMargin.Value = 1;
            ctrl.Margin.TopMargin.Value = 2;
            ctrl.Margin.RightMargin.Value = 3;
            ctrl.Margin.BottomMargin.Value = 4;

            ctrl.TextAnchor = eTextAnchoringType.Distributed;
            ctrl.TextAlignment = eTextAlignment.Right;

            ctrl.LayoutFlow = eLayoutFlow.VerticalIdeographic;
            ctrl.Orientation = eShapeOrientation.TopToBottom;
            ctrl.ReadingOrder = eReadingOrder.LeftToRight;
            ctrl.AutomaticSize = true;
            
            Assert.AreEqual(eEditAs.Absolute ,ctrl.EditAs);
            Assert.AreEqual("A1", ctrl.FmlaTxbx.Address);
            
            Assert.IsFalse(ctrl.Margin.Automatic);
            Assert.AreEqual(1, ctrl.Margin.LeftMargin.Value);
            Assert.AreEqual(eMeasurementUnits.Millimeters, ctrl.Margin.LeftMargin.Unit);
            Assert.AreEqual(2, ctrl.Margin.TopMargin.Value);
            Assert.AreEqual(eMeasurementUnits.Millimeters, ctrl.Margin.TopMargin.Unit);
            Assert.AreEqual(3, ctrl.Margin.RightMargin.Value);
            Assert.AreEqual(eMeasurementUnits.Millimeters, ctrl.Margin.RightMargin.Unit);
            Assert.AreEqual(4, ctrl.Margin.BottomMargin.Value);
            Assert.AreEqual(eMeasurementUnits.Millimeters, ctrl.Margin.BottomMargin.Unit);

            Assert.IsTrue(ctrl.AutomaticSize);

            Assert.AreEqual(eTextAnchoringType.Distributed, ctrl.TextAnchor);
            Assert.AreEqual(eTextAlignment.Right, ctrl.TextAlignment);

            Assert.AreEqual(eLayoutFlow.VerticalIdeographic, ctrl.LayoutFlow);
            Assert.AreEqual(eShapeOrientation.TopToBottom, ctrl.Orientation);
            Assert.AreEqual(eReadingOrder.LeftToRight, ctrl.ReadingOrder);            
        }
        [TestMethod]
        public void AddCheckboxTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("Checkbox");
            var ctrl = _ws.Drawings.AddCheckBoxControl("Checkbox 1");
            ctrl.Macro = "Checkbox_Click";
            ctrl.Fill.Style = eVmlFillType.Gradient;
            ctrl.Fill.SecondColor.ColorString= "#ff8200";
            ctrl.Fill.GradientSettings.Focus = 100;
            ctrl.Fill.GradientSettings.Angle = -135;
            ctrl.Fill.Color.ColorString = "#000082";
            ctrl.Fill.GradientSettings.SetGradientColors(new VmlGradiantColor(0, Color.Red), new VmlGradiantColor(50, Color.Orange), new VmlGradiantColor(100, Color.Yellow));            
            ctrl.Fill.Opacity = 97;
            ctrl.Fill.Recolor = true;
            ctrl.Fill.SecondColorOpacity = 50;
            ctrl.Border.LineStyle = eVmlLineStyle.ThickThin;
            ctrl.Border.Width.Value = 1;
            ctrl.Border.Width.Unit = eMeasurementUnits.Pixels;
            ctrl.LinkedCell = _ws.Cells["F1"];
            ctrl.Checked = eCheckState.Mixed;
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 100);
            
            var codeModule = _pck.Workbook.VbaProject.Modules.AddModule("CheckboxCode");
            _codeModule.Code += "Sub Checkbox_Click()\r\n  MsgBox \"Clicked Checkbox!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddRadioButtonTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("RadioButton");
            var groupBox = _ws.Drawings.AddGroupBoxControl("Groupbox 1");
            groupBox.SetPosition(80, 80);
            groupBox.SetSize(240, 120);

            var ctrl = _ws.Drawings.AddRadioButtonControl("Option Button 1");
            ctrl.Macro = "RadioButton_Click";
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 30);

            var ctrl2 = _ws.Drawings.AddControl("RadioButton 2", eControlType.RadioButton);
            ctrl2.Macro = "RadioButton_Click";
            ctrl2.SetPosition(130, 100);
            ctrl2.SetSize(200, 30);

            var ctrl3 = _ws.Drawings.AddControl("RadioButton 3", eControlType.RadioButton);
            ctrl3.Macro = "RadioButton_Click";
            ctrl3.SetPosition(160, 100);
            ctrl3.SetSize(200, 30);

            var groupBox2 = _ws.Drawings.AddControl("Groupbox 2", eControlType.GroupBox);
            groupBox2.SetPosition(780, 80);
            groupBox2.SetSize(240, 120);

            var ctrl4 = _ws.Drawings.AddControl("RadioButton 4", eControlType.RadioButton).As.Control.RadioButton;
            ctrl4.FirstButton = true;
            ctrl4.SetPosition(800, 100);
            ctrl4.SetSize(200, 30);

            var ctrl5 = _ws.Drawings.AddControl("RadioButton 5", eControlType.RadioButton);
            ctrl5.SetPosition(830, 100);
            ctrl5.SetSize(200, 30);

            var ctrl6 = _ws.Drawings.AddControl("RadioButton 6", eControlType.RadioButton);
            ctrl6.SetPosition(860, 100);
            ctrl6.SetSize(200, 30);

            var codeModule = _pck.Workbook.VbaProject.Modules.AddModule("RadioButtonCode");
            _codeModule.Code += "Sub RadioButton_Click()\r\n  MsgBox \"Clicked RadioButton!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddCheckboxWithFrameFillTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("CheckboxWithImageFill");
            var ctrl = _ws.Drawings.AddCheckBoxControl("Checkbox 2");
            ctrl.Fill.Style = eVmlFillType.Frame;
            ctrl.Fill.PatternPictureSettings.Image.SetImage(Properties.Resources.Test1JpgByteArray, ePictureType.Jpg);
            ctrl.Fill.PatternPictureSettings.AspectRatio = eVmlAspectRatio.AtLeast;
            ctrl.Fill.PatternPictureSettings.Size = "0,0";
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 100);
        }
        [TestMethod]
        public void AddCheckboxWithTileFillTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("CheckboxWithTileFill");
            var ctrl = _ws.Drawings.AddCheckBoxControl("Checkbox 2");
            ctrl.Fill.Style = eVmlFillType.Tile;
            ctrl.Fill.PatternPictureSettings.Image.SetImage(Properties.Resources.CodeTif, ePictureType.Tif);
            ctrl.Fill.Color.SetColor(Color.Black);
            ctrl.Fill.Recolor = true;
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 100);
        }
        [TestMethod]
        public void AddCheckboxWithPatternFillTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("CheckboxWithPatternFill");
            var ctrl = _ws.Drawings.AddCheckBoxControl("Checkbox 2");
            ctrl.Fill.Style = eVmlFillType.Pattern;
            ctrl.Fill.PatternPictureSettings.Image.SetImage(Properties.Resources.VmlPatternImage, ePictureType.Png);
            ctrl.Fill.Color.SetColor(Color.Red);
            ctrl.Fill.SecondColor.SetColor(Color.Yellow);
            ctrl.Fill.Recolor = true;
            ctrl.AutoFill = false;
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 100);
        }

        [TestMethod]
        public void AddDropDownTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("DropDown");
            var ctrl = _ws.Drawings.AddDropDownControl("DropDown 1");
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
            var ctrl = _ws.Drawings.AddListBoxControl("ListBox 1");
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
            var ctrl = _ws.Drawings.AddLabelControl("Label 1");
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
            var ctrl = _ws.Drawings.AddSpinButtonControl("SpinButton 1");
            ctrl.Macro = "SpinButton_Click";
            ctrl.SetPosition(500, 100);
            ctrl.SetSize(200, 100);

            _ws.Cells["G1"].Value = 3;

            ctrl.LinkedCell = _ws.Cells["G1"];

            _codeModule.Code += "Sub SpinButton_Click()\r\n  MsgBox \"Selected SpinButton!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddScrollbarButtonTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("Scrollbar");
            var ctrl = _ws.Drawings.AddScrollBarControl("Scrollbar 1");
            ctrl.Macro = "ScrolbarButton_Click";
            ctrl.SetPosition(500, 100);
            //ctrl.SetSize(200, 100);

            _ws.Cells["G1"].Value = 3;

            ctrl.LinkedCell = _ws.Cells["G1"];
            ctrl.MaxValue = 100;
            ctrl.MinValue = 2;
            ctrl.Page = 20;
            _codeModule.Code += "Sub SpinButton_Click()\r\n  MsgBox \"Selected SpinButton!!\"\r\nEnd Sub\r\n";
        }

        [TestMethod]
        public void AddGroupBoxTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("GroupBox");
            var ctrl = _ws.Drawings.AddGroupBoxControl("GroupBox 1");
            ctrl.Macro = "GroupBox_Click";
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

            r3.UpdateXml();
            Assert.AreEqual(r3.From.Row, r3._control.From.Row);
            Assert.AreEqual(r3.From.RowOff, r3._control.From.RowOff);
            Assert.AreEqual(r3.From.Column, r3._control.From.Column);
            Assert.AreEqual(r3.From.ColumnOff, r3._control.From.ColumnOff);

            Assert.AreEqual(r3.To.Row, r3._control.To.Row);
            Assert.AreEqual(r3.To.RowOff, r3._control.To.RowOff);
            Assert.AreEqual(r3.To.Column, r3._control.To.Column);
            Assert.AreEqual(r3.To.ColumnOff, r3._control.To.ColumnOff);

            ctrl.Group(r1, r2, r3);

            _codeModule.Code += "Sub GroupBox_Click()\r\n  MsgBox \"Clicked GroupBox!!\"\r\nEnd Sub\r\n";
        }
        [TestMethod]
        public void AddControlHeaderAndComment()
        {
            var ws = _pck.Workbook.Worksheets.Add("HeaderControlAndComment");
            ws.HeaderFooter.OddHeader.CenteredText = "Before ";
            var img = ws.HeaderFooter.OddHeader.InsertPicture(Properties.Resources.Test1, PictureAlignment.Centered);
            img.Title = "Renamed Image";

            ws.Comments.Add(ws.Cells["A1"], "Comment in cell A1", "JK");
            var btn = ws.Drawings.AddButtonControl("Button 1");
            btn.SetPosition(100, 100);
        }
        [TestMethod]
        public void ValidateVmlGetColor()
        {
            Assert.AreEqual(Color.FromArgb(0xFF,0x33, 0x99, 0x66).ToArgb(), ExcelVmlDrawingColor.GetColor("#396 [57]").ToArgb());
            Assert.AreEqual(Color.FromArgb(0xFF, 0xFF, 0xCC, 0x99).ToArgb(), ExcelVmlDrawingColor.GetColor("#fc9").ToArgb());
            Assert.AreEqual(Color.FromArgb(0xFF, 0x00, 0x00, 0x82).ToArgb(), ExcelVmlDrawingColor.GetColor("#000082").ToArgb());
            Assert.AreEqual(Color.Red.ToArgb(), ExcelVmlDrawingColor.GetColor("Red").ToArgb());
            Assert.AreEqual(Color.Blue.ToArgb(), (long)ExcelVmlDrawingColor.GetColor("Blue [0]").ToArgb());
            Assert.AreEqual(Color.FromArgb(0xFF, 200, 100, 0).ToArgb(), ExcelVmlDrawingColor.GetColor("rgb (200, 100, 0)").ToArgb()); 
        }
        [TestMethod]
        public void RemoveControlTest()
        {
            _ws = _pck.Workbook.Worksheets.Add("RemoveControl");
            var ctrl = _ws.Drawings.AddGroupBoxControl("GroupBox 1");
            Assert.AreEqual(1, _ws.Drawings.Count);
            //_ws.Drawings.Remove(ctrl);
            //Assert.AreEqual(0, _ws.Drawings.Count);
        }
    }
}
