using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Control
{
    [TestClass]
    public class ReadControlTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenTemplatePackage("control.xlsm");
            _ws = _pck.Workbook.Worksheets[0];
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
            _pck.Dispose();
        }
        [TestMethod]
        public void ValidateNumberOfDrawings()
        {
            Assert.AreEqual(11, _ws.Drawings.Count);
        }
        [TestMethod]
        public void ValidateButtonControl()
        {
            /**** Button ****/
            Assert.IsInstanceOfType(_ws.Drawings[0], typeof(ExcelControlButton));
            var button = _ws.Drawings[0].As.Control.Button;
            Assert.AreEqual(eControlType.Button, button.ControlType);
            Assert.IsTrue(button.LockedText);
            Assert.AreEqual("Button 1", button.Name);
            Assert.AreEqual("[0]!Button1_Click", button.Macro);
            Assert.IsFalse(button.AutoPict);
            Assert.IsFalse(button.AutoFill);
            Assert.IsFalse(button.DefaultSize);
        }
        [TestMethod]
        public void ValidateDropDownControl()
        {
            /**** DropDown ****/
            Assert.IsInstanceOfType(_ws.Drawings[1], typeof(ExcelControlDropDown));
            var dropDown = _ws.Drawings[1].As.Control.DropDown;
            Assert.AreEqual(eControlType.DropDown, dropDown.ControlType);
            Assert.AreEqual("Drop Down 3", dropDown.Name);
            Assert.AreEqual("[0]!DropDown3_Change", dropDown.Macro);
            Assert.IsFalse(dropDown.AutoPict);
            Assert.IsTrue(dropDown.AutoFill);
            Assert.IsFalse(dropDown.DefaultSize);
            Assert.AreEqual(0, dropDown.SelectedIndex);
            Assert.AreEqual(eDropStyle.Combo, dropDown.DropStyle);
            Assert.AreEqual(8, dropDown.DropLines);
            Assert.AreEqual("$A$1", dropDown.LinkedCell.Address);
            Assert.AreEqual("$K$4:$K$8", dropDown.InputRange.Address);

        }
        [TestMethod]
        public void ValidateLabelControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[2], typeof(ExcelControlLabel));
            Assert.AreEqual(eControlType.Label, _ws.Drawings[2].As.Control.Label.ControlType);
            var label = _ws.Drawings[2].As.Control.Label;
            Assert.AreEqual("Label 6", label.Name);
            Assert.AreEqual("Label 6", label.Text);
            Assert.IsTrue(label.LockedText);
        }
        [TestMethod]
        public void ValidateListboxControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[3], typeof(ExcelControlListBox));
            var listBox = _ws.Drawings[3].As.Control.ListBox;
            Assert.AreEqual("$J$4:$K$8", listBox.InputRange.Address);
            Assert.AreEqual("$A$1:$A$2", listBox.LinkedCell.Address);
            Assert.AreEqual(0, listBox.SelectedIndex);
            Assert.AreEqual(eSelectionType.Extended, listBox.SelectionType);
            Assert.AreEqual(2, listBox.MultiSelection.Length); ;
            Assert.AreEqual(3, listBox.MultiSelection[0]); ;
            Assert.AreEqual(2, listBox.MultiSelection[1]); ;
            Assert.AreEqual("List Box 7", listBox.Name);
        }
        [TestMethod]
        public void ValidateCheckboxControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[4], typeof(ExcelControlCheckBox));
            var checkbox = _ws.Drawings[4].As.Control.CheckBox;
            Assert.AreEqual(eControlType.CheckBox, checkbox.ControlType);
            Assert.AreEqual(eCheckState.Checked, checkbox.Checked);
            Assert.IsTrue(checkbox.LockedText);
            Assert.IsFalse(checkbox.ThreeDEffects);
            Assert.AreEqual("Check Box 9", checkbox.Name);
            Assert.AreEqual("Check Box 9", checkbox.Text);
        }
        [TestMethod]
        public void ValidateCheckboxWithTileControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[8], typeof(ExcelControlCheckBox));
            var checkbox = _ws.Drawings[8].As.Control.CheckBox;
            Assert.AreEqual(eControlType.CheckBox, checkbox.ControlType);
            Assert.AreEqual(eCheckState.Checked, checkbox.Checked);
            Assert.AreEqual("Check Box 12", checkbox.Name);
            Assert.AreEqual("Check Box 12", checkbox.Text);
            Assert.AreEqual(OfficeOpenXml.Drawing.Vml.eVmlFillType.Tile, checkbox.Fill.Style);
            Assert.IsNotNull(checkbox.Fill.PatternPictureSettings.Image);

        }
        [TestMethod]
        public void ValidateCheckboxWithFrameControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[9], typeof(ExcelControlCheckBox));
            var checkbox = _ws.Drawings[9].As.Control.CheckBox;
            Assert.AreEqual(eControlType.CheckBox, checkbox.ControlType);
            Assert.AreEqual(eCheckState.Checked, checkbox.Checked);
            Assert.AreEqual("Check Box 13", checkbox.Name);
            Assert.AreEqual("Check Box 13", checkbox.Text);
            Assert.AreEqual(OfficeOpenXml.Drawing.Vml.eVmlFillType.Frame, checkbox.Fill.Style);
            Assert.IsNotNull(checkbox.Fill.PatternPictureSettings.Image);
        }
        [TestMethod]
        public void ValidateCheckboxWithPatternControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[10], typeof(ExcelControlCheckBox));
            var checkbox = _ws.Drawings[10].As.Control.CheckBox;
            Assert.AreEqual(eControlType.CheckBox, checkbox.ControlType);
            Assert.AreEqual(eCheckState.Checked, checkbox.Checked);
            Assert.AreEqual("Check Box 14", checkbox.Name);
            Assert.AreEqual("Check Box 14", checkbox.Text);
            Assert.AreEqual(OfficeOpenXml.Drawing.Vml.eVmlFillType.Pattern, checkbox.Fill.Style);
            Assert.IsNotNull(checkbox.Fill.PatternPictureSettings.Image);
        }

        [TestMethod]
        public void ValidateSpinbuttonControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[5], typeof(ExcelControlSpinButton));
            var spin = _ws.Drawings[5].As.Control.Spin;
            Assert.AreEqual(eControlType.SpinButton, spin.ControlType);
            Assert.AreEqual("$K$22", spin.LinkedCell.Address);
            Assert.AreEqual(3, spin.Increment);
            Assert.AreEqual(0, spin.MinValue);
            Assert.AreEqual(30000, spin.MaxValue);
            Assert.AreEqual(18, spin.Value);
            Assert.AreEqual("Spinner 10", spin.Name);
        }
        [TestMethod]
        public void ValidateGroupBoxControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[6], typeof(ExcelControlGroupBox));
            Assert.AreEqual(eControlType.GroupBox, _ws.Drawings[6].As.Control.GroupBox.ControlType);
            var groupBox = _ws.Drawings[6].As.Control.GroupBox;
            Assert.AreEqual("[0]!GroupBox5_Click", groupBox.Macro);
            Assert.AreEqual("Group Box 5", groupBox.Name);
            Assert.AreEqual("Group Box 5", groupBox.Text);
        }        
        [TestMethod]
        public void ValidateRadioButtonControl()
        {
            Assert.IsInstanceOfType(_ws.Drawings[7], typeof(ExcelControlRadioButton));
            Assert.AreEqual(eControlType.RadioButton, _ws.Drawings[7].As.Control.RadioButton.ControlType);
            var radioButton = _ws.Drawings[7].As.Control.RadioButton;
            Assert.IsTrue(radioButton.LockedText);
            Assert.IsTrue(radioButton.FirstButton);
            Assert.IsTrue(radioButton.Checked);
            Assert.IsFalse(radioButton.ThreeDEffects);

            Assert.AreEqual("Option Button 11", radioButton.Name);
            Assert.AreEqual("Option Button 11", radioButton.Text);
        }
        [TestMethod]
        public void ValidateDrawingGroup()
        {
            var ws = _pck.Workbook.Worksheets[1];

            Assert.IsInstanceOfType(ws.Drawings[0], typeof(ExcelGroupShape));
            Assert.AreEqual(2, ws.Drawings.Count);
            var grp = (ExcelGroupShape)ws.Drawings[0];
            Assert.AreEqual(4, grp.Drawings.Count);
            Assert.AreEqual(eDrawingType.Control, grp.Drawings[0].DrawingType);
            Assert.AreEqual(2028825, grp.Drawings[0].Size.Width);

            //grp.Drawings.Clear();
        }
        [TestMethod]
        public void ValidateDrawingUnGroup()
        {
            var ws = _pck.Workbook.Worksheets[1];

            Assert.IsInstanceOfType(ws.Drawings[0], typeof(ExcelGroupShape));
            Assert.AreEqual(2, ws.Drawings.Count);
            var grp = (ExcelGroupShape)ws.Drawings[0];
            Assert.AreEqual(4, grp.Drawings.Count);
            Assert.AreEqual(eDrawingType.Control, grp.Drawings[0].DrawingType);
            Assert.AreEqual(2028825, grp.Drawings[0].Size.Width);
        }

    }
}
