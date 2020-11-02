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
            _pck = OpenPackage("FormControl.xlsx",true);
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
            _ws.Drawings.AddControl("button 1", eControlType.Button);
        }
    }
}
