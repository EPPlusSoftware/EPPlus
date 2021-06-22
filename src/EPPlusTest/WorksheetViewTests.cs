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
        public void SplitPanges()
        {
            using (var p = OpenTemplatePackage("SplitRead.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.IsNotNull(ws.View);
                Assert.IsNotNull(ws.View.PaneSettings);
                Assert.AreEqual(ePaneState.Frozen, ws.View.PaneSettings.State);
                Assert.AreEqual(4, ws.View.Panes.Length);
            }
        }
    }
}
