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
            var ws = _pck.Workbook.Worksheets.Add("TableStyle");
            var s=_pck.Workbook.Styles.CreateSlicerStyle("CustomSlicerStyle1");
            s.WholeTable.Style.Font.Color.SetColor(Color.LightGray);
            s.HeaderRow.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);

            s.SelectedItemWithData.Style.Border.Top.Style = ExcelBorderStyle.Dotted;
            s.SelectedItemWithData.Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;
            s.SelectedItemWithData.Style.Border.Bottom.Color.SetColor(Color.Black);
            s.SelectedItemWithData.Style.Border.Left.Style = ExcelBorderStyle.Dotted;
            s.SelectedItemWithData.Style.Border.Right.Style = ExcelBorderStyle.Dotted;

            LoadTestdata(ws);
            var tbl=ws.Tables.Add(ws.Cells["A1:D101"], "Table1");
            var slicer = tbl.Columns[0].AddSlicer();
            slicer.SetPosition(100, 100);
            slicer.StyleName = "CustomSlicerStyle1";

            //Assert
            //Assert.AreEqual(ExcelFillStyle.Solid, s.FirstRowStripe.Style.Fill.PatternType);
            //Assert.AreEqual(Color.Red.ToArgb(), s.WholeTable.Style.Font.Color.Color.Value.ToArgb());
            //Assert.AreEqual(Color.LightBlue.ToArgb(), s.FirstRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
            //Assert.AreEqual(ExcelFillStyle.Solid, s.SecondRowStripe.Style.Fill.PatternType);
            //Assert.AreEqual(Color.LightYellow.ToArgb(), s.SecondRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        }
        //[TestMethod]
        //public void ReadTableStyle()
        //{
        //    using (var p = OpenTemplatePackage("TableStyleRead.xlsx"))
        //    {
        //        Assert.AreEqual(3, p.Workbook.Styles.TableStyles.Count);
        //        var s = p.Workbook.Styles.TableStyles["CustomTableStyle1"].As.TableStyle;                
        //        Assert.AreEqual("CustomTableStyle1", s.Name);

        //        //Assert
        //        Assert.AreEqual(ExcelFillStyle.Solid, s.FirstRowStripe.Style.Fill.PatternType);
        //        Assert.AreEqual(Color.Red.ToArgb(), s.WholeTable.Style.Font.Color.Color.Value.ToArgb());
        //        Assert.AreEqual(Color.LightBlue.ToArgb(), s.FirstRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        //        Assert.AreEqual(ExcelFillStyle.Solid, s.SecondRowStripe.Style.Fill.PatternType);
        //        Assert.AreEqual(Color.LightYellow.ToArgb(), s.SecondRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        //    }
        //}
        //[TestMethod]
        //public void AddPivotTableStyle()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("PivotTableStyle");
        //    var s = _pck.Workbook.Styles.CreatePivotTableStyle("CustomPivotTableStyle1");
        //    s.WholeTable.Style.Font.Color.SetColor(Color.DarkBlue);
        //    s.FirstRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //    s.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
        //    s.FirstRowStripe.BandSize = 2;
        //    s.SecondRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //    s.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
        //    s.SecondRowStripe.BandSize = 3;
        //    LoadTestdata(ws);
        //    var pt = ws.PivotTables.Add(ws.Cells["G2"], ws.Cells["A1:D101"], "PivotTable1");

        //    pt.StyleName = "CustomPivotTableStyle1";
        //}
        //[TestMethod]
        //public void ReadPivotTableStyle()
        //{
        //    using (var p = OpenTemplatePackage("TableStyleRead.xlsx"))
        //    {
        //        Assert.AreEqual(3, p.Workbook.Styles.TableStyles.Count);
        //        var s = p.Workbook.Styles.TableStyles["CustomPivotTableStyle1"].As.PivotTableStyle;
        //        Assert.AreEqual("CustomPivotTableStyle1", s.Name);

        //        //Assert
        //        Assert.AreEqual(Color.DarkBlue.ToArgb(), s.WholeTable.Style.Font.Color.Color.Value.ToArgb());

        //        Assert.AreEqual(ExcelFillStyle.Solid, s.FirstRowStripe.Style.Fill.PatternType);
        //        Assert.AreEqual(Color.LightGreen.ToArgb(), s.FirstRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        //        Assert.AreEqual(2, s.FirstRowStripe.BandSize);

        //        Assert.AreEqual(ExcelFillStyle.Solid, s.SecondRowStripe.Style.Fill.PatternType);
        //        Assert.AreEqual(Color.LightGray.ToArgb(), s.SecondRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        //        Assert.AreEqual(3, s.SecondRowStripe.BandSize);
        //    }
        //}
        //[TestMethod]
        //public void AddTableAndPivotTableStyle()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("SharedTableStyle");
        //    var s = _pck.Workbook.Styles.CreateTableAndPivotTableStyle("CustomTableAndPivotTableStyle1");
        //    s.WholeTable.Style.Font.Color.SetColor(Color.DarkMagenta);

        //    s.FirstColumnStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;            
        //    s.FirstColumnStripe.Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
        //    s.FirstColumnStripe.BandSize = 2;

        //    s.SecondColumnStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //    s.SecondColumnStripe.Style.Fill.BackgroundColor.SetColor(Color.LightPink);
        //    s.SecondColumnStripe.BandSize = 2;

        //    LoadTestdata(ws);
        //    var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table2");
        //    tbl.StyleName = "CustomTableAndPivotTableStyle1";
            
        //    var pt = ws.PivotTables.Add(ws.Cells["G2"], tbl, "PivotTable2");
        //    pt.RowFields.Add(pt.Fields[0]);
        //    pt.DataFields.Add(pt.Fields[3]);
        //    pt.ShowRowStripes = true;
        //    pt.StyleName = "CustomTableAndPivotTableStyle1";
        //}
        //[TestMethod]
        //public void ReadTableAndPivotTableStyle()
        //{
        //    using (var p = OpenTemplatePackage("TableStyleRead.xlsx"))
        //    {
        //        Assert.AreEqual(3, p.Workbook.Styles.TableStyles.Count);
        //        var s = p.Workbook.Styles.TableStyles["CustomTableAndPivotTableStyle1"].As.TableAndPivotTableStyle;
        //        Assert.AreEqual("CustomTableAndPivotTableStyle1", s.Name);

        //        //Assert
        //        Assert.AreEqual(Color.DarkMagenta.ToArgb(), s.WholeTable.Style.Font.Color.Color.Value.ToArgb());

        //        Assert.AreEqual(ExcelFillStyle.Solid, s.FirstColumnStripe.Style.Fill.PatternType);
        //        Assert.AreEqual(Color.LightCyan.ToArgb(), s.FirstColumnStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        //        Assert.AreEqual(2, s.FirstColumnStripe.BandSize);

        //        Assert.AreEqual(ExcelFillStyle.Solid, s.SecondColumnStripe.Style.Fill.PatternType);
        //        Assert.AreEqual(Color.LightPink.ToArgb(), s.SecondColumnStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        //        Assert.AreEqual(2, s.SecondColumnStripe.BandSize);
        //    }
        //}

        //[TestMethod]
        //public void AlterTableStyle()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("TableRowStyle");
        //    LoadTestdata(ws);
        //    var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table3");
        //    var ns = _pck.Workbook.Styles.CreateNamedStyle("TableCellStyle2");
        //    ns.Style.Font.Color.SetColor(Color.Red);
        //    //var s = _pck.Workbook.Styles.CreateTableStyle("CustomTableStyle1");
        //    //s.HeaderRow.Style.Font.Color.SetColor(Color.Red);
        //    //s.FirstRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //    //s.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
        //    //s.SecondRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //    //s.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);


        //    tbl.TableStyle = OfficeOpenXml.Table.TableStyles.None;
        //    //tbl.HeaderRowStyleName = "TableCellStyle2";
        //    tbl.Range.Offset(0, 0, 1, tbl.Range.Columns).StyleName= "TableCellStyle2";
        //    ////tbl.StyleName = "CustomTableStyle1";
        //    ////tbl.HeaderRowStyle.Border.Bottom.Style=ExcelBorderStyle.Dashed;
        //    ////tbl.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);
        //    //tbl.DataStyle.Font.Color.SetColor(Color.Red);
        //    //tbl.Columns[0].DataStyle.Font.Color.SetColor(Color.Green);
        //}

        //[TestMethod]
        //public void CopyTableRowStyle()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("CopyTableRowStyleSource");
        //    LoadTestdata(ws);
        //    var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table4");

        //    tbl.HeaderRowStyle.Border.Bottom.Style=ExcelBorderStyle.Dashed;
        //    tbl.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);
        //    tbl.DataStyle.Font.Color.SetColor(Color.Red);
        //    tbl.Columns[0].DataStyle.Font.Color.SetColor(Color.Green);

        //    var wsCopy = _pck.Workbook.Worksheets.Add("CopyTableRowStyleCopy", ws);

        //    Assert.AreEqual(ExcelBorderStyle.Dashed, tbl.HeaderRowStyle.Border.Bottom.Style);
        //    Assert.AreEqual(Color.Black.ToArgb(),tbl.HeaderRowStyle.Border.Bottom.Color.Color.Value.ToArgb());
        //    Assert.AreEqual(Color.Red.ToArgb(), tbl.DataStyle.Font.Color.Color.Value.ToArgb());
        //    Assert.AreEqual(Color.Green.ToArgb(), tbl.Columns[0].DataStyle.Font.Color.Color.Value.ToArgb());

        //}
        //[TestMethod]
        //public void CopyTableRowStyleNewPackage()
        //{
        //    using (var p1 = new ExcelPackage())
        //    {
        //        var ws = p1.Workbook.Worksheets.Add("CopyTableRowStyleSource");
        //        LoadTestdata(ws);
        //        var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table4");

        //        tbl.HeaderRowStyle.Border.Bottom.Style = ExcelBorderStyle.Dashed;
        //        tbl.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);
        //        tbl.DataStyle.Font.Color.SetColor(Color.Red);
        //        tbl.Columns[0].DataStyle.Font.Color.SetColor(Color.Green);

        //        using (var p2 = new ExcelPackage())
        //        {
        //            var wsCopy = p2.Workbook.Worksheets.Add("CopyTableRowStyleCopy", ws);

        //            Assert.AreEqual(ExcelBorderStyle.Dashed, tbl.HeaderRowStyle.Border.Bottom.Style);
        //            Assert.AreEqual(Color.Black.ToArgb(), tbl.HeaderRowStyle.Border.Bottom.Color.Color.Value.ToArgb());
        //            Assert.AreEqual(Color.Red.ToArgb(), tbl.DataStyle.Font.Color.Color.Value.ToArgb());
        //            Assert.AreEqual(Color.Green.ToArgb(), tbl.Columns[0].DataStyle.Font.Color.Color.Value.ToArgb());
        //            SaveWorkbook("TableDxfCopy.xlsx", p2);
        //        }
        //    }
        //}
    }
}


