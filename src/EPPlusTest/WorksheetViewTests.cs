/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/31/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class WorksheetViewTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("WorksheetView.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void SetTopLeftCellToH15()
        {
            var ws = _pck.Workbook.Worksheets.Add("TopLeftH15");
            ws.View.TopLeftCell = "H15";

            Assert.AreEqual("H15", ws.View.TopLeftCell);
        }
        [TestMethod]
        public void SetTopLeftCellToNullAfterItHasBeenSet()
        {
            var ws = _pck.Workbook.Worksheets.Add("TopLeftBlank");
            ws.View.TopLeftCell = "H15";

            Assert.AreEqual("H15", ws.View.TopLeftCell);
            ws.View.TopLeftCell = null;

            Assert.AreEqual("", ws.View.TopLeftCell);
        }
        [TestMethod]
        public void ReadFrozenPanes()
        {
            using (var p = OpenTemplatePackage("FrozenRead.xlsx"))
            {
                //Worksheet 1
                var ws = p.Workbook.Worksheets[0];
                Assert.IsNotNull(ws.View);
                Assert.IsNotNull(ws.View.PaneSettings);
                Assert.AreEqual(ePaneState.Frozen, ws.View.PaneSettings.State);
                Assert.AreEqual(4, ws.View.Panes.Length);
                Assert.AreEqual("H56" ,ws.View.PaneSettings.TopLeftCell);
                Assert.AreEqual(7D, ws.View.PaneSettings.YSplit);
                Assert.AreEqual(7D, ws.View.PaneSettings.XSplit);

                //Worksheet 2
                ws = p.Workbook.Worksheets[1];
                Assert.IsNotNull(ws.View);
                Assert.IsNotNull(ws.View.PaneSettings);
                Assert.AreEqual(ePaneState.Frozen, ws.View.PaneSettings.State);
                Assert.AreEqual(ePanePosition.BottomRight, ws.View.PaneSettings.ActivePanePosition);
                Assert.AreEqual(3, ws.View.Panes.Length);
                Assert.AreEqual("D8", ws.View.PaneSettings.TopLeftCell);
                Assert.AreEqual(7D, ws.View.PaneSettings.YSplit);
                Assert.AreEqual(3D, ws.View.PaneSettings.XSplit);
            }
        }
        [TestMethod]
        public void ReadSplitPanes()
        {
            using (var p = OpenTemplatePackage("SplitPanes.xlsx"))
            {
                //Worksheet 1
                var ws = p.Workbook.Worksheets[0];
                Assert.IsNotNull(ws.View);
                Assert.IsNotNull(ws.View.PaneSettings);
                Assert.AreEqual(ePaneState.Split, ws.View.PaneSettings.State);
                Assert.AreEqual(4, ws.View.Panes.Length);

                Assert.AreEqual(4230, ws.View.PaneSettings.XSplit);
                Assert.AreEqual(3300, ws.View.PaneSettings.YSplit);
                Assert.AreEqual(3300, ws.Column(1).Width);
            }
        }
        [TestMethod]
        public void SplitPanesBoth()
        {
            var ws = _pck.Workbook.Worksheets.Add("SplitPanes");
            ws.View.TopLeftCell = "G200";
            ws.View.SplitPanes(2, 2);
            ws.View.ActiveCell = "B2";
        }
        [TestMethod]
        public void SplitPanesRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("SplitPanesRow");
            ws.View.TopLeftCell = "A200";
            ws.View.SplitPanes(2, 1);
            ws.View.ActiveCell = "A201";
            ws.View.Panes[0].ActiveCell = "A5";
            ws.View.PaneSettings.TopLeftCell = "A182";
        }

        [TestMethod]
        public void SplitPanesColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("SplitPanesColumn");
            ws.View.TopLeftCell = "A200";
            ws.View.SplitPanes(0, 2);
            ws.View.ActiveCell = "A201";
            ws.View.Panes[0].ActiveCell = "A5";
        }

        [TestMethod]
        public void SplitPanesNormal()
        {
            var ws = _pck.Workbook.Worksheets.Add("SplitPanesNormal48");
            //_pck.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
            _pck.Workbook.Styles.NamedStyles[0].Style.Font.Size = 48;
            ws.View.TopLeftCell = "G200";
            ws.View.SplitPanes(3, 3);
            ws.View.TopLeftPane.ActiveCell = "B2";
        }
        [TestMethod]
        public void SplitPanesNormal11Ariel()
        {
            var ws = _pck.Workbook.Worksheets.Add("SplitPanesNormal48RH");
            _pck.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
            _pck.Workbook.Styles.NamedStyles[0].Style.Font.Size = 11;
            ws.View.TopLeftCell = "G2";
            ws.View.SplitPanes(3, 3);
            ws.View.ActiveCell = "B2";
        }
    }
}
