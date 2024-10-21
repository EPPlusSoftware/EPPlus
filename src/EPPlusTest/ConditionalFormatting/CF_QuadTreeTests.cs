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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml.Style;

namespace EPPlusTest.ConditionalFormatting
{
    /// <summary>
    /// Test the Conditional Formatting feature
    /// </summary>
    [TestClass]
    public class CF_QuadTree : TestBase
    {
        private static ExcelPackage _pck;

        [ClassInitialize()]
        public static void Init(TestContext testContext)
        {
            _pck = OpenPackage("CFQuadTree.xlsx", true);
        }
        [TestMethod]
        public void QuadTreeRemoveCF()
        {
            var ws = _pck.Workbook.Worksheets.Add("RemoveSingleCF");
            
            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B1:C2");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType=ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("B1:C2");
            rule2.Text = "Z";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.ConditionalFormatting.Remove(rule1);
            formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();

            Assert.AreEqual(1, formats.Count);
            Assert.AreEqual(formats[0], rule2);

            ws.ConditionalFormatting.Remove(rule2);
            formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();

            Assert.AreEqual(0, formats.Count);
        }
        [TestMethod]
        public void QuadTreeDeleteRowValidate()
        {
            var ws = _pck.Workbook.Worksheets.Add("DeleteRowCF");

            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B2:C3");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("A1:D4");
            rule2.Text = "C";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var rule3 = ws.ConditionalFormatting.AddEndsWith("B10:C20");
            rule2.Text = "Z";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.DeleteRow(2, 2);
            formats = ws.Cells["C2:C8"].ConditionalFormatting.GetConditionalFormattings();

            Assert.AreEqual(2, formats.Count);
            Assert.AreEqual(formats[0], rule2);
            Assert.AreEqual("A1:D2", formats[0].Address.Address);
            Assert.AreEqual("B8:C18", formats[1].Address.Address);
        }
        [TestMethod]
        public void QuadTreeDeleteColumnValidate()
        {
            var ws = _pck.Workbook.Worksheets.Add("DeleteColumnCF");

            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B2:C3");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("A1:D4");
            rule2.Text = "C";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var rule3 = ws.ConditionalFormatting.AddEndsWith("AC1:AE3");
            rule2.Text = "Z";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.DeleteColumn(2, 2);
            formats = ws.Cells["B2:AA3"].ConditionalFormatting.GetConditionalFormattings();

            Assert.AreEqual(2, formats.Count);
            Assert.AreEqual(formats[0], rule2);
            Assert.AreEqual("A1:B4", formats[0].Address.Address);
            Assert.AreEqual("AA1:AC3", formats[1].Address.Address);
        }
        [TestMethod]
        public void QuadTreeInsertRowValidate()
        {
            var ws = _pck.Workbook.Worksheets.Add("InsertRowCF");

            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B2:C3");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("A1:D4");
            rule2.Text = "C";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var rule3 = ws.ConditionalFormatting.AddEndsWith("AC1:AE3");
            rule2.Text = "Z";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.InsertRow(5, 2);
            formats = ws.Cells["C4"].ConditionalFormatting.GetConditionalFormattings();

            Assert.AreEqual(1, formats.Count);
            Assert.AreEqual(formats[0], rule2);
            Assert.AreEqual("A1:D6", formats[0].Address.Address);
        }
        [TestMethod]
        public void QuadTreeInsertColumnValidate()
        {
            var ws = _pck.Workbook.Worksheets.Add("InsertColumnCF");

            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B2:C3");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("A1:D4");
            rule2.Text = "C";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var rule3 = ws.ConditionalFormatting.AddEndsWith("AA1:AC10");
            rule3.Text = "Z";
            rule3.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule3.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule3.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.InsertColumn(4, 2);
            formats = ws.Cells["C4:AC4"].ConditionalFormatting.GetConditionalFormattings();

            Assert.AreEqual(2, formats.Count);
            Assert.AreEqual(formats[0], rule2);
            Assert.AreEqual("A1:F4", formats[0].Address.Address);
            Assert.AreEqual("AC1:AE10", formats[1].Address.Address);
        }
        [TestMethod]
        public void QuadTreeDeleteShiftLeftValidate()
        {
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeLeftCF");

            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B2:C3");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("A1:D4");
            rule2.Text = "C";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.Cells["C2"].Delete(eShiftTypeDelete.Left);
            formats = ws.Cells["C2"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(1, formats.Count);
            Assert.AreEqual("A1:D1,A2:C2,A3:D4", formats[0].Address.Address);
            
            formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);
        }
        [TestMethod]
        public void QuadTreeDeleteShiftUpValidate()
        {
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeUpCF");

            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B2:C3");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("A1:D4");
            rule2.Text = "C";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.Cells["C2"].Delete(eShiftTypeDelete.Up);
            formats = ws.Cells["C2"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);
            Assert.AreEqual("B2:B3,C2", formats[0].Address.Address);
            Assert.AreEqual("A1:B4,C1:C3,D1:D4", formats[1].Address.Address);

            formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

        }
        [TestMethod]
        public void QuadTreeInsertShiftRightValidate()
        {
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeLeftCF");

            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B2:C3");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("A1:D4");
            rule2.Text = "C";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.Cells["C2"].Insert(eShiftTypeInsert.Right);
            formats = ws.Cells["C2"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);
            Assert.AreEqual("B2:D2,B3:C3", formats[0].Address.Address);
            Assert.AreEqual("A1:D1,A2:E2,A3:D4", formats[1].Address.Address);
            formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);
        }
        [TestMethod]
        public void QuadTreeInsertShiftDownValidate()
        {
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeLeftCF");

            var rule1 = ws.ConditionalFormatting.AddBeginsWith("B2:C3");
            rule1.Text = "B";
            rule1.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule1.Style.Fill.BackgroundColor.SetColor(Color.Red);

            var rule2 = ws.ConditionalFormatting.AddEndsWith("A1:D4");
            rule2.Text = "C";
            rule2.Style.Fill.Style = eDxfFillStyle.PatternFill;
            rule2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule2.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);

            ws.Cells["C2"].Insert(eShiftTypeInsert.Down);
            formats = ws.Cells["C2"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);
            Assert.AreEqual("B2:B3,C2:C4", formats[0].Address.Address);
            Assert.AreEqual("A1:B4,C1:C5,D1:D4", formats[1].Address.Address);
            formats = ws.Cells["C2:C3"].ConditionalFormatting.GetConditionalFormattings();
            Assert.AreEqual(2, formats.Count);
        }

    }
}