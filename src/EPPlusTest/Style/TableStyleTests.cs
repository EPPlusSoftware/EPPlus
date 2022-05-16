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
using OfficeOpenXml.Table;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;

namespace EPPlusTest.Style
{
    [TestClass]
    public class TableStyleTests : TestBase
    {
        static ExcelPackage _pck;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("TableStyle.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
            if (File.Exists(fileName))
            {
                File.Copy(fileName, dirName + "\\TableStyleRead.xlsx", true);
            }
        }
        [TestMethod]
        public void AddTableStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableStyle");
            var s=_pck.Workbook.Styles.CreateTableStyle("CustomTableStyle1");
            s.WholeTable.Style.Font.Color.SetColor(Color.Red);
            s.FirstRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            s.SecondRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

            LoadTestdata(ws);
            var tbl=ws.Tables.Add(ws.Cells["A1:D101"], "Table1");            
            tbl.StyleName = "CustomTableStyle1";            
            //Assert
            Assert.AreEqual(ExcelFillStyle.Solid, s.FirstRowStripe.Style.Fill.PatternType);
            Assert.AreEqual(Color.Red.ToArgb(), s.WholeTable.Style.Font.Color.Color.Value.ToArgb());
            Assert.AreEqual(Color.LightBlue.ToArgb(), s.FirstRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
            Assert.AreEqual(ExcelFillStyle.Solid, s.SecondRowStripe.Style.Fill.PatternType);
            Assert.AreEqual(Color.LightYellow.ToArgb(), s.SecondRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        }
        [TestMethod]
        public void AddTableStyleFromTemplate()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableStyleFromTempl");
            var s = _pck.Workbook.Styles.CreateTableStyle("CustomTableStyleFromTempl1", OfficeOpenXml.Table.TableStyles.Medium5);
            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table2");
            tbl.StyleName = "CustomTableStyleFromTempl1";

            //Assert
            Assert.AreEqual(eThemeSchemeColor.Text1, s.WholeTable.Style.Font.Color.Theme);
            Assert.IsTrue(s.HeaderRow.Style.Font.Bold.Value);
            Assert.AreEqual(ExcelBorderStyle.Double, s.TotalRow.Style.Border.Top.Style);
            Assert.AreEqual(ExcelFillStyle.Solid, s.FirstRowStripe.Style.Fill.PatternType);
            Assert.AreEqual(0.79998D, Math.Round(s.FirstColumnStripe.Style.Fill.BackgroundColor.Tint.Value,5));
        }

        [TestMethod]
        public void AddTableStyleFromOtherStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableStyleFromOther");
            var sc = _pck.Workbook.Styles.CreateTableStyle("CustomTableStyleFromTemplCopy", OfficeOpenXml.Table.TableStyles.Light14);
            var s = _pck.Workbook.Styles.CreateTableStyle("CustomTableStyleFromOther1", sc);
            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableOtherStyle");
            tbl.StyleName = "CustomTableStyleFromOther1";

            //Assert
            Assert.AreEqual(eThemeSchemeColor.Text1, s.WholeTable.Style.Font.Color.Theme);
            Assert.IsTrue(s.HeaderRow.Style.Font.Bold.Value);
            Assert.AreEqual(ExcelBorderStyle.Double, s.TotalRow.Style.Border.Top.Style);
        }
    
        [TestMethod]
        public void ReadTableStyle()
        {
            using (var p = OpenTemplatePackage("TableStyleRead.xlsx"))
            {
                var s = p.Workbook.Styles.TableStyles["CustomTableStyle1"];
                if (s == null) Assert.Inconclusive("CustomTableStyle1 does not exists in workbook");
                Assert.IsNotNull(s.As.TableStyle);
                Assert.AreEqual("CustomTableStyle1", s.Name);

                //Assert
                Assert.AreEqual(ExcelFillStyle.Solid, s.FirstRowStripe.Style.Fill.PatternType);
                Assert.AreEqual(Color.Red.ToArgb(), s.WholeTable.Style.Font.Color.Color.Value.ToArgb());
                Assert.AreEqual(Color.LightBlue.ToArgb(), s.FirstRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
                Assert.AreEqual(ExcelFillStyle.Solid, s.SecondRowStripe.Style.Fill.PatternType);
                Assert.AreEqual(Color.LightYellow.ToArgb(), s.SecondRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
            }
        }
        [TestMethod]
        public void AddPivotTableStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotTableStyle");
            var s = _pck.Workbook.Styles.CreatePivotTableStyle("CustomPivotTableStyle1");
            s.WholeTable.Style.Font.Color.SetColor(Color.DarkBlue);
            s.FirstRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            s.FirstRowStripe.BandSize = 2;
            s.SecondRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            s.SecondRowStripe.BandSize = 3;
            LoadTestdata(ws);
            var pt = ws.PivotTables.Add(ws.Cells["G2"], ws.Cells["A1:D101"], "PivotTable1");

            pt.StyleName = "CustomPivotTableStyle1";
        }
        [TestMethod]
        public void AddPivotTableStyleFromTemplate()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotTableStyleFromTempl");
            var s = _pck.Workbook.Styles.CreatePivotTableStyle("CustomPivotTableStyleFromTempl1", PivotTableStyles.Dark2);

            LoadTestdata(ws);
            var pt = ws.PivotTables.Add(ws.Cells["G2"], ws.Cells["A1:D101"], "PivotTable2");
            pt.ColumnFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.StyleName = "CustomPivotTableStyleFromTempl1";
        }

        [TestMethod]
        public void ReadPivotTableStyle()
        {
            using (var p = OpenTemplatePackage("TableStyleRead.xlsx"))
            {
                var s = p.Workbook.Styles.TableStyles["CustomPivotTableStyle1"];
                if (s == null) Assert.Inconclusive("CustomPivotTableStyle1 does not exists in workbook");
                var ps = s.As.PivotTableStyle;

                Assert.AreEqual("CustomPivotTableStyle1", ps.Name);

                //Assert
                Assert.AreEqual(Color.DarkBlue.ToArgb(), ps.WholeTable.Style.Font.Color.Color.Value.ToArgb());

                Assert.AreEqual(ExcelFillStyle.Solid, ps.FirstRowStripe.Style.Fill.PatternType);
                Assert.AreEqual(Color.LightGreen.ToArgb(), ps.FirstRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
                Assert.AreEqual(2, ps.FirstRowStripe.BandSize);

                Assert.AreEqual(ExcelFillStyle.Solid, ps.SecondRowStripe.Style.Fill.PatternType);
                Assert.AreEqual(Color.LightGray.ToArgb(), ps.SecondRowStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
                Assert.AreEqual(3, ps.SecondRowStripe.BandSize);
            }
        }
        [TestMethod]
        public void AddTableAndPivotTableStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("SharedTableStyle");
            var s = _pck.Workbook.Styles.CreateTableAndPivotTableStyle("CustomTableAndPivotTableStyle1");
            if (s == null) Assert.Inconclusive("CustomTableAndPivotTableStyle1 does not exists in workbook");

            s.WholeTable.Style.Font.Color.SetColor(Color.DarkMagenta);

            s.FirstColumnStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;            
            s.FirstColumnStripe.Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
            s.FirstColumnStripe.BandSize = 2;

            s.SecondColumnStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.SecondColumnStripe.Style.Fill.BackgroundColor.SetColor(Color.LightPink);
            s.SecondColumnStripe.BandSize = 2;

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table3");
            tbl.StyleName = "CustomTableAndPivotTableStyle1";
            
            var pt = ws.PivotTables.Add(ws.Cells["G2"], tbl, "PivotTable3");
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.ShowRowStripes = true;
            pt.StyleName = "CustomTableAndPivotTableStyle1";
        }
        [TestMethod]
        public void ReadTableAndPivotTableStyle()
        {
            using (var p = OpenTemplatePackage("TableStyleRead.xlsx"))
            {
                var s = p.Workbook.Styles.TableStyles["CustomTableAndPivotTableStyle1"];
                if (s == null) Assert.Inconclusive("CustomTableAndPivotTableStyle1 style is not present in the workbook");
                var tpts =s.As.TableAndPivotTableStyle;
                Assert.AreEqual("CustomTableAndPivotTableStyle1", tpts.Name);

                //Assert
                Assert.AreEqual(Color.DarkMagenta.ToArgb(), tpts.WholeTable.Style.Font.Color.Color.Value.ToArgb());

                Assert.AreEqual(ExcelFillStyle.Solid, tpts.FirstColumnStripe.Style.Fill.PatternType);
                Assert.AreEqual(Color.LightCyan.ToArgb(), tpts.FirstColumnStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
                Assert.AreEqual(2, tpts.FirstColumnStripe.BandSize);

                Assert.AreEqual(ExcelFillStyle.Solid, tpts.SecondColumnStripe.Style.Fill.PatternType);
                Assert.AreEqual(Color.LightPink.ToArgb(), tpts.SecondColumnStripe.Style.Fill.BackgroundColor.Color.Value.ToArgb());
                Assert.AreEqual(2, tpts.SecondColumnStripe.BandSize);
            }
        }

        [TestMethod]
        public void AlterTableStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableRowStyle");
            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table4");
            var ns = _pck.Workbook.Styles.CreateNamedStyle("TableCellStyle2");
            ns.Style.Font.Color.SetColor(Color.Red);


            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.None;
            tbl.Range.Offset(0, 0, 1, tbl.Range.Columns).StyleName= "TableCellStyle2";
        }

        [TestMethod]
        public void CopyTableRowStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("CopyTableRowStyleSource");
            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table5");

            tbl.HeaderRowStyle.Border.Bottom.Style=ExcelBorderStyle.Dashed;
            tbl.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);
            tbl.HeaderRowStyle.Border.Top.Style = ExcelBorderStyle.Dashed;
            tbl.HeaderRowStyle.Border.Top.Color.SetColor(Color.Black);

            tbl.DataStyle.Font.Color.SetColor(Color.Red);
            tbl.Columns[0].DataStyle.Font.Color.SetColor(Color.Green);
            tbl.TotalsRowStyle.Font.Size = 15;
            tbl.TotalsRowStyle.Font.Color.SetColor(eThemeSchemeColor.Accent5);
            tbl.ShowTotal = true;
            tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            //tbl.Columns[1].TotalsRowStyle.NumberFormat.Format = "#,##0.00";

            var wsCopy = _pck.Workbook.Worksheets.Add("CopyTableRowStyleCopy", ws);

            Assert.AreEqual(ExcelBorderStyle.Dashed, tbl.HeaderRowStyle.Border.Bottom.Style);
            Assert.AreEqual(Color.Black.ToArgb(),tbl.HeaderRowStyle.Border.Bottom.Color.Color.Value.ToArgb());
            Assert.AreEqual(Color.Red.ToArgb(), tbl.DataStyle.Font.Color.Color.Value.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), tbl.Columns[0].DataStyle.Font.Color.Color.Value.ToArgb());

        }
        [TestMethod]
        public void CopyTableRowStyleNewPackage()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("CopyTableRowStyleSource");
                LoadTestdata(ws);
                var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table6");
                tbl.HeaderRowStyle.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                tbl.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);
                tbl.DataStyle.Font.Color.SetColor(Color.Red);
                tbl.Columns[0].DataStyle.Font.Color.SetColor(Color.Green);
                tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
                tbl.TotalsRowStyle.Font.Size = 15;
                tbl.TotalsRowStyle.Font.Color.SetColor(eThemeSchemeColor.Accent5);
                tbl.ShowTotal = true;
                tbl.Columns[1].DataStyle.NumberFormat.Format = "#,##0.00";
                tbl.Columns[1].TotalsRowStyle.NumberFormat.Format = "#,##0.00";

                using (var p2 = new ExcelPackage())
                {
                    var wsCopy = p2.Workbook.Worksheets.Add("CopyTableRowStyleCopy", ws);
                    var tblCopy = wsCopy.Tables[0];
                    Assert.AreEqual(ExcelBorderStyle.Dashed, tbl.HeaderRowStyle.Border.Bottom.Style);
                    Assert.AreEqual(Color.Black.ToArgb(), tbl.HeaderRowStyle.Border.Bottom.Color.Color.Value.ToArgb());
                    Assert.AreEqual(Color.Red.ToArgb(), tbl.DataStyle.Font.Color.Color.Value.ToArgb());
                    Assert.AreEqual(Color.Green.ToArgb(), tbl.Columns[0].DataStyle.Font.Color.Color.Value.ToArgb());

                    Assert.AreEqual(ExcelBorderStyle.Dashed, tblCopy.HeaderRowStyle.Border.Bottom.Style);
                    Assert.AreEqual(Color.Black.ToArgb(), tblCopy.HeaderRowStyle.Border.Bottom.Color.Color.Value.ToArgb());
                    Assert.AreEqual(Color.Red.ToArgb(), tblCopy.DataStyle.Font.Color.Color.Value.ToArgb());
                    Assert.AreEqual(Color.Green.ToArgb(), tblCopy.Columns[0].DataStyle.Font.Color.Color.Value.ToArgb());
                    Assert.AreEqual(tbl.Name, tblCopy.Name); //Keep name if this table does not exist in the workbook.

                    SaveWorkbook("TableDxfCopy.xlsx", p2);
                }
            }
        }     
        [TestMethod]
        public void SetStyleWhenAddingRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("SetStyleWhenAddingRow");
            LoadTestdata(ws);

            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table7");

            tbl.DataStyle.Font.Italic=true;
            tbl.AddRow(2);

            Assert.IsTrue(ws.Cells["A101"].Style.Font.Italic);
            Assert.IsTrue(ws.Cells["D102"].Style.Font.Italic);
        }
        [TestMethod]
        public void SetStyleWhenInsertingRowFirst()
        {
            var ws = _pck.Workbook.Worksheets.Add("SetStyleWhenInsertingRowFirst");
            LoadTestdata(ws);

            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table8");

            tbl.ShowHeader = false; 
            tbl.DataStyle.Font.Strike = true;
            tbl.InsertRow(0, 3);

            Assert.IsTrue(ws.Cells["A2"].Style.Font.Strike);
            Assert.IsTrue(ws.Cells["D3"].Style.Font.Strike);
        }
        [TestMethod]
        public void SetStyleWhenInsertingColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("SetStyleWhenInsertingColumn");
            LoadTestdata(ws);

            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table9");

            tbl.ShowHeader = true;
            tbl.ShowTotal = true;
            tbl.HeaderRowStyle.Border.Bottom.Style = ExcelBorderStyle.DashDotDot;
            tbl.HeaderRowStyle.Border.Bottom.Color.Auto = true;
            tbl.DataStyle.Font.Italic = true;
            tbl.TotalsRowStyle.Border.Top.Style = ExcelBorderStyle.Dashed;
            tbl.TotalsRowStyle.Border.Top.Color.Theme = eThemeSchemeColor.Accent6;
            tbl.Columns.Insert(2,3);

            Assert.IsTrue(ws.Cells["C1"].Style.Border.Bottom.Color.Auto);
            Assert.AreEqual(ExcelBorderStyle.DashDotDot,ws.Cells["D1"].Style.Border.Bottom.Style);
            Assert.IsTrue(ws.Cells["C50"].Style.Font.Italic);
            Assert.IsTrue(ws.Cells["D100"].Style.Font.Italic);
            Assert.AreEqual(ExcelBorderStyle.Dashed, tbl.TotalsRowStyle.Border.Top.Style);
            Assert.AreEqual(eThemeSchemeColor.Accent6, tbl.TotalsRowStyle.Border.Top.Color.Theme);
        }
        [TestMethod]
        public void SetStyleWhenInsertingColumnFirst()
        {
            var ws = _pck.Workbook.Worksheets.Add("SetStyleWhenInsertingColumnFirst");
            LoadTestdata(ws);

            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table10");
            
            tbl.DataStyle.Font.Strike = true;
            tbl.InsertRow(0, 3);

            Assert.IsTrue(ws.Cells["A2"].Style.Font.Strike);
            Assert.IsTrue(ws.Cells["D3"].Style.Font.Strike);
        }
        [TestMethod]
        public void SetStyleWhenInsertingColumnFirstNoHeader()
        {
            var ws = _pck.Workbook.Worksheets.Add("SetStyleInsertingColumnFirstNH");
            LoadTestdata(ws);

            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table11");

            tbl.DataStyle.Font.Strike = true;
            tbl.ShowHeader = false;
            tbl.InsertRow(0, 3);            
            Assert.IsTrue(ws.Cells["A1"].Style.Font.Strike);
            Assert.IsTrue(ws.Cells["D3"].Style.Font.Strike);
        }
    }
}