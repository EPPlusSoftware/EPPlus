using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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
    public class ControlTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            //_pck = OpenPackage("DrawingBorder.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            //var dirName = _pck.File.DirectoryName;
            //var fileName = _pck.File.FullName;

            //SaveAndCleanup(_pck);
            //File.Copy(fileName, dirName + "\\DrawingBorderRead.xlsx", true);
        }

        [TestMethod]
        public void ReadControls()
        {
            using (var p = OpenTemplatePackage("control.xlsm"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual(8, ws.Drawings.Count);

                /**** Button ****/
                Assert.IsInstanceOfType(ws.Drawings[0], typeof(ExcelControlButton));
                var button = ws.Drawings[0].As.Control.Button;
                Assert.AreEqual(eControlType.Button, button.ControlType);
                Assert.IsTrue(button.LockedText);
                Assert.AreEqual("Button 1", button.Name);
                Assert.AreEqual("[0]!Button1_Click", button.Macro);                
                Assert.IsFalse(button.AutoPict);
                Assert.IsFalse(button.AutoFill);
                Assert.IsFalse(button.DefaultSize);

                /**** DropDown ****/
                Assert.IsInstanceOfType(ws.Drawings[1], typeof(ExcelControlDropDown));
                var dropDown = ws.Drawings[1].As.Control.DropDown;
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

                Assert.IsInstanceOfType(ws.Drawings[2], typeof(ExcelControlLabel));
                Assert.AreEqual(eControlType.Label, ws.Drawings[2].As.Control.Label.ControlType);
                var label = ws.Drawings[2].As.Control.Label;
                Assert.AreEqual("Label 6", label.Name);
                Assert.AreEqual("Label 6", label.Text);

                Assert.IsInstanceOfType(ws.Drawings[3], typeof(ExcelControlListBox));
                var listBox = ws.Drawings[3].As.Control.ListBox;
                var range = listBox.InputRange;
                var linkedCell = listBox.LinkedCell;

                Assert.IsInstanceOfType(ws.Drawings[4], typeof(ExcelControlCheckBox));
                var checkbox = ws.Drawings[4].As.Control.CheckBox;
                Assert.AreEqual(eControlType.CheckBox, checkbox.ControlType);

                Assert.IsInstanceOfType(ws.Drawings[5], typeof(ExcelControlSpin));
                var spin = ws.Drawings[5].As.Control.Spin;
                Assert.AreEqual(eControlType.Spin, spin.ControlType);


                Assert.IsInstanceOfType(ws.Drawings[6], typeof(ExcelControlGroupBox));
                Assert.AreEqual(eControlType.GroupBox, ws.Drawings[6].As.Control.GroupBox.ControlType);
                var groupBox = ws.Drawings[6].As.Control.GroupBox;
                Assert.AreEqual("Group Box 5", groupBox.Name);
                Assert.AreEqual("Group Box 5", groupBox.Text);

                Assert.IsInstanceOfType(ws.Drawings[7], typeof(ExcelControlRadioButton));
                Assert.AreEqual(eControlType.RadioButton, ws.Drawings[7].As.Control.RadioButton.ControlType);
                var radioButton = ws.Drawings[7].As.Control.RadioButton;

                Assert.AreEqual("Option Button 11", radioButton.Name);
                Assert.AreEqual("Option Button 11", radioButton.Text);
            }
        }
    }
}
