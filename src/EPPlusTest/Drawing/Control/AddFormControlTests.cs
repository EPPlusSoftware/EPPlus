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
    public class AddControlTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("FormControl.xlsm",true);
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
            var ctrl = _ws.Drawings.AddControl("Button 1", eControlType.Button);
            ctrl.Macro = "Button1_Click";
            ctrl.SetPosition(100, 100);
            ctrl.SetSize(200, 100);

            _pck.Workbook.CreateVBAProject();
            var codeModule = _pck.Workbook.VbaProject.Modules.AddModule("ButtonCode");
            codeModule.Code= "Sub Button1_Click()\r\n  MsgBox \"Clicked!!\"\r\nEnd Sub";
        }
    }
}
