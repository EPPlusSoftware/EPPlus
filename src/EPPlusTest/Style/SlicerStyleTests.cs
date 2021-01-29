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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;

namespace EPPlusTest.Style
{
    [TestClass]
    public class SlicerStyleTests : TestBase
    {
        static ExcelPackage _pck;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("SlicerStyle.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
            if (File.Exists(fileName))
            {
                File.Copy(fileName, dirName + "\\SlicerStyleRead.xlsx", true);
            }
        }
        [TestMethod]
        public void AddSlicerStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("SlicerStyleAdd");
            var s=_pck.Workbook.Styles.CreateSlicerStyle("CustomSlicerStyle1");
            s.WholeTable.Style.Font.Color.SetColor(Color.LightGray);
            s.HeaderRow.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);

            s.SelectedItemWithData.Style.Font.Bold=true;
            s.SelectedItemWithData.Style.Border.Top.Style = ExcelBorderStyle.Dotted;
            s.SelectedItemWithData.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
            s.SelectedItemWithData.Style.Border.Bottom.Color.SetColor(Color.Green);
            s.SelectedItemWithData.Style.Border.Left.Style = ExcelBorderStyle.DashDotDot;
            s.SelectedItemWithData.Style.Border.Right.Style = ExcelBorderStyle.None;
            s.HoveredSelectedItemWithData.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent4);
            s.HoveredUnselectedItemWithData.Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);

            LoadTestdata(ws);
            var tbl=ws.Tables.Add(ws.Cells["A1:D101"], "Table1");
            var slicer = tbl.Columns[0].AddSlicer();
            slicer.SetPosition(100, 100);
            slicer.StyleName = "CustomSlicerStyle1";
            
            //Assert
            Assert.AreEqual("CustomSlicerStyle1", slicer.StyleName);
            Assert.AreEqual(Color.LightGray.ToArgb(), s.WholeTable.Style.Font.Color.Color.Value.ToArgb());
            Assert.AreEqual(Color.DarkGray.ToArgb(), s.HeaderRow.Style.Fill.BackgroundColor.Color.Value.ToArgb());

            Assert.IsTrue(s.SelectedItemWithData.Style.Font.Bold.Value);
            Assert.AreEqual(ExcelBorderStyle.Dotted, s.SelectedItemWithData.Style.Border.Top.Style);
            Assert.AreEqual(ExcelBorderStyle.Hair, s.SelectedItemWithData.Style.Border.Bottom.Style);
            Assert.AreEqual(ExcelBorderStyle.DashDotDot, s.SelectedItemWithData.Style.Border.Left.Style);
            Assert.AreEqual(ExcelBorderStyle.None, s.SelectedItemWithData.Style.Border.Right.Style);
            Assert.AreEqual(eThemeSchemeColor.Accent4, s.HoveredSelectedItemWithData.Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(Color.LightGoldenrodYellow.ToArgb(), s.HoveredUnselectedItemWithData.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        }
        [TestMethod]
        public void AddSlicerStyleFromTemplate()
        {
            var ws = _pck.Workbook.Worksheets.Add("SlicerStyleTemplate");
            var s = _pck.Workbook.Styles.CreateSlicerStyle("CustomSlicerStyleFromTemplate", eSlicerStyle.Dark1);

            s.WholeTable.Style.Font.Name = "Arial";
            s.HeaderRow.Style.Font.Italic = true;

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table2");
            var slicer = tbl.Columns[0].AddSlicer();
            slicer.SetPosition(100, 100);
            slicer.StyleName = "CustomSlicerStyleFromTemplate";

            //Assert
            Assert.AreEqual(eDxfFillStyle.GradientFill, s.HoveredSelectedItemWithData.Style.Fill.Style);
            Assert.AreEqual(2, s.HoveredSelectedItemWithData.Style.Fill.Gradient.Colors.Count);
            Assert.AreEqual(0, s.HoveredSelectedItemWithData.Style.Fill.Gradient.Colors[0].Position);
            Assert.AreEqual(Color.FromArgb(0xFF, 0XF8, 0XE1, 0X62), s.HoveredSelectedItemWithData.Style.Fill.Gradient.Colors[0].Color.Color);
            Assert.AreEqual(Color.FromArgb(0xFF, 0XFC, 0XF7, 0XE0), s.HoveredSelectedItemWithData.Style.Fill.Gradient.Colors[1].Color.Color);
            Assert.AreEqual(100, s.HoveredSelectedItemWithData.Style.Fill.Gradient.Colors[1].Position);
        }
        [TestMethod]
        public void AddSlicerStyleFromOther()
        {
            var ws = _pck.Workbook.Worksheets.Add("SlicerStyleCopyOther");
            var s = _pck.Workbook.Styles.CreateSlicerStyle("CustomSlicerStyleToCopy", eSlicerStyle.Other2);

            var sc= _pck.Workbook.Styles.CreateSlicerStyle("CustomSlicerStyleCopy", s);

            sc.SelectedItemWithData.Style.Fill.Style = eDxfFillStyle.PatternFill;
            sc.SelectedItemWithData.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Background2);
            sc.SelectedItemWithData.Style.Fill.PatternType = ExcelFillStyle.LightGray;

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table3");
            var slicer = tbl.Columns[0].AddSlicer();
            slicer.SetPosition(100, 100);
            slicer.StyleName = "CustomSlicerStyleCopy";

            Assert.AreEqual(eThemeSchemeColor.Background2, sc.SelectedItemWithData.Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(ExcelFillStyle.LightGray, sc.SelectedItemWithData.Style.Fill.PatternType);
        }

        [TestMethod]
        public void AddSlicerStyleFromOtherNewPackage()
        {
            var ws = _pck.Workbook.Worksheets.Add("SlicerStyleCopyOtherPck");
            var s = _pck.Workbook.Styles.CreateSlicerStyle("CustomSlicerStyleToCopyOther", eSlicerStyle.Other2);

            var fmt = "#,##0.0";
            s.HoveredUnselectedItemWithNoData.Style.NumberFormat.Format = fmt;
            s.WholeTable.Style.Font.Name = "Arial";

            using (var p = new ExcelPackage())
            {
                var sc = p.Workbook.Styles.CreateSlicerStyle("CustomSlicerStyleCopyPck", s);
                ws=p.Workbook.Worksheets.Add("CopiedSlicerStyle");
                LoadTestdata(ws);
                var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table3");
                var slicer = tbl.Columns[0].AddSlicer();
                slicer.SetPosition(100, 100);
                slicer.StyleName = "CustomSlicerStyleCopyPck";

                Assert.AreEqual(fmt, sc.HoveredUnselectedItemWithNoData.Style.NumberFormat.Format);

                SaveWorkbook("SlicerStyleNewPackage.Xlsx", p);
            }
        }


        [TestMethod]
        public void ReadSlicerStyle()
        {
            using (var p = OpenTemplatePackage("SlicerStyleRead.xlsx"))
            {
                var s = p.Workbook.Styles.SlicerStyles["CustomSlicerStyle1"];
                if (s == null) Assert.Inconclusive("Custom style does not exists");

                Assert.AreEqual("CustomSlicerStyle1", s.Name);

                //Assert
                Assert.AreEqual(Color.LightGray.ToArgb(), s.WholeTable.Style.Font.Color.Color.Value.ToArgb());
                Assert.AreEqual(Color.DarkGray.ToArgb(), s.HeaderRow.Style.Fill.BackgroundColor.Color.Value.ToArgb());

                Assert.IsTrue(s.SelectedItemWithData.Style.Font.Bold.Value);
                Assert.AreEqual(ExcelBorderStyle.Dotted, s.SelectedItemWithData.Style.Border.Top.Style);
                Assert.AreEqual(ExcelBorderStyle.Hair, s.SelectedItemWithData.Style.Border.Bottom.Style);
                Assert.AreEqual(ExcelBorderStyle.DashDotDot, s.SelectedItemWithData.Style.Border.Left.Style);
                Assert.AreEqual(ExcelBorderStyle.None, s.SelectedItemWithData.Style.Border.Right.Style);
                Assert.AreEqual(eThemeSchemeColor.Accent4, s.HoveredSelectedItemWithData.Style.Fill.BackgroundColor.Theme);
                Assert.AreEqual(Color.LightGoldenrodYellow.ToArgb(), s.HoveredUnselectedItemWithData.Style.Fill.BackgroundColor.Color.Value.ToArgb());
            }
        }
    }
}


