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
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Core
{
    [TestClass]
    public class WorksheetHiddenTests : TestBase
    {
        [TestMethod]
        public void HideTest_0Based()
        {
            using (var pck = new ExcelPackage())
            {
                pck.Compatibility.IsWorksheets1Based = false;
                var ws = pck.Workbook.Worksheets.Add("Hidden");
                pck.Workbook.Worksheets.Add("Visible");
                ws.Cells["A1"].Value = "This workbook is hidden";
                ws.Hidden = eWorkSheetHidden.Hidden;
                Assert.AreEqual(eWorkSheetHidden.Hidden, ws.Hidden);
                Assert.AreEqual(1, pck.Workbook.View.ActiveTab);
                SaveWorkbook("HiddenSecondWorbook.xlsx", pck);
            }
        }

        [TestMethod]
        public void VeryHideTest_0Based()
        {
            using (var pck = new ExcelPackage())
            {
                pck.Compatibility.IsWorksheets1Based = false;
                var ws = pck.Workbook.Worksheets.Add("VeryHidden");
                pck.Workbook.Worksheets.Add("Visible");
                ws.Cells["A1"].Value = "This worksheet is veryhidden";
                ws.Hidden = eWorkSheetHidden.VeryHidden;
                Assert.AreEqual(eWorkSheetHidden.VeryHidden, ws.Hidden);
                Assert.AreEqual(1, pck.Workbook.View.ActiveTab);
                SaveWorkbook("VeryHiddenSecondWorbook.xlsx", pck);
            }
        }
        [TestMethod]
        public void HideTest_1Based()
        {
            using (var pck = new ExcelPackage())
            {
                pck.Compatibility.IsWorksheets1Based = true;
                var ws = pck.Workbook.Worksheets.Add("Hidden");
                pck.Workbook.Worksheets.Add("Visible");
                ws.Cells["A1"].Value = "This workbook is hidden";
                ws.Hidden = eWorkSheetHidden.Hidden;
                Assert.AreEqual(eWorkSheetHidden.Hidden, ws.Hidden);
                Assert.AreEqual(1, pck.Workbook.View.ActiveTab);
                SaveWorkbook("HiddenSecondWorbook.xlsx", pck);
            }
        }

        [TestMethod]
        public void VeryHideTest_1Based()
        {
            using (var pck = new ExcelPackage())
            {
                pck.Compatibility.IsWorksheets1Based = true;
                var ws = pck.Workbook.Worksheets.Add("VeryHidden");
                pck.Workbook.Worksheets.Add("Visible");
                ws.Cells["A1"].Value = "This worksheet is veryhidden";
                ws.Hidden = eWorkSheetHidden.VeryHidden;
                Assert.AreEqual(eWorkSheetHidden.VeryHidden, ws.Hidden);
                Assert.AreEqual(1, pck.Workbook.View.ActiveTab);
                SaveWorkbook("VeryHiddenSecondWorbook.xlsx", pck);
            }
        }
    }
}
