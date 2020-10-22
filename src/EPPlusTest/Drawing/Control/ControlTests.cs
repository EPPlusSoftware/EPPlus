using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                Assert.AreEqual(5, ws.Drawings.Count);

                Assert.IsInstanceOfType(ws.Drawings[0], typeof(ExcelControlButton));
                Assert.AreEqual(eControlType.Button, ws.Drawings[0].As.Control.Button.ControlType);

                Assert.IsInstanceOfType(ws.Drawings[1], typeof(ExcelControlDropDown));
                Assert.AreEqual(eControlType.DropDown, ws.Drawings[1].As.Control.DropDown.ControlType);

                Assert.IsInstanceOfType(ws.Drawings[2], typeof(ExcelControlGroupBox));
                Assert.AreEqual(eControlType.GroupBox, ws.Drawings[2].As.Control.GroupBox.ControlType);

                Assert.IsInstanceOfType(ws.Drawings[3], typeof(ExcelControlLabel));
                Assert.AreEqual(eControlType.Label, ws.Drawings[3].As.Control.Label.ControlType);

                Assert.IsInstanceOfType(ws.Drawings[4], typeof(ExcelControlListBox));
                var range = ws.Drawings[4].As.Control.ListBox.InputRange;
                var linkedCell = ws.Drawings[4].As.Control.ListBox.LinkedCell;



            }
        }
    }
}
