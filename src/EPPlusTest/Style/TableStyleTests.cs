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
                File.Copy(fileName, dirName + "\\PivotTableNamedStylesRead.xlsx", true);
            }
        }
        [TestMethod]
        public void VerifyColumnStyle()
        {
            using (var p = OpenTemplatePackage("PivotTableNamedStyles.xlsx"))
            {
                Assert.AreEqual(2, p.Workbook.Styles.TableStyles.Count);
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
        }
        [TestMethod]
        public void AddPivotTableStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotTableStyle");
            var s = _pck.Workbook.Styles.CreatePivotTableStyle("CustomPivotTableStyle1");
            s.WholeTable.Style.Font.Color.SetColor(Color.DarkBlue);
            s.FirstRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            s.SecondRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["G2"], ws.Cells["A1:D101"], "PivotTable1");
            tbl.StyleName = "CustomPivotTableStyle1";
        }

        [TestMethod]
        public void AddTableAndPivotTableStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("SharedTableStyle");
            var s = _pck.Workbook.Styles.CreateTableAndPivotTableStyle("CustomTableAndPivotTableStyle1");
            s.WholeTable.Style.Font.Color.SetColor(Color.DarkMagenta);

            s.FirstRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;            
            s.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
            s.FirstRowStripe.BandSize = 2;

            s.SecondRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightPink);
            s.SecondRowStripe.BandSize = 2;

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table2");
            tbl.StyleName = "CustomTableAndPivotTableStyle1";
            
            var pt = ws.PivotTables.Add(ws.Cells["G2"], tbl, "PivotTable2");
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.ShowRowStripes = true;
            pt.StyleName = "CustomTableAndPivotTableStyle1";
        }
        [TestMethod]
        public void AlterTableStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableRowStyle");
            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table3");
            var ns = _pck.Workbook.Styles.CreateNamedStyle("TableCellStyle2");
            ns.Style.Font.Color.SetColor(Color.Red);
            //var s = _pck.Workbook.Styles.CreateTableStyle("CustomTableStyle1");
            //s.HeaderRow.Style.Font.Color.SetColor(Color.Red);
            //s.FirstRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            //s.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            //s.SecondRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            //s.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);


            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.None;
            //tbl.HeaderRowStyleName = "TableCellStyle2";
            tbl.Range.Offset(0, 0, 1, tbl.Range.Columns).StyleName= "TableCellStyle2";
            ////tbl.StyleName = "CustomTableStyle1";
            ////tbl.HeaderRowStyle.Border.Bottom.Style=ExcelBorderStyle.Dashed;
            ////tbl.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);
            //tbl.DataStyle.Font.Color.SetColor(Color.Red);
            //tbl.Columns[0].DataStyle.Font.Color.SetColor(Color.Green);
        }

        [TestMethod]
        public void CopyTableRowStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("CopyTableRowStyleSource");
            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table4");

            tbl.HeaderRowStyle.Border.Bottom.Style=ExcelBorderStyle.Dashed;
            tbl.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);
            tbl.DataStyle.Font.Color.SetColor(Color.Red);
            tbl.Columns[0].DataStyle.Font.Color.SetColor(Color.Green);

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
                var tbl = ws.Tables.Add(ws.Cells["A1:D101"], "Table4");

                tbl.HeaderRowStyle.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                tbl.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);
                tbl.DataStyle.Font.Color.SetColor(Color.Red);
                tbl.Columns[0].DataStyle.Font.Color.SetColor(Color.Green);

                using (var p2 = new ExcelPackage())
                {
                    var wsCopy = p2.Workbook.Worksheets.Add("CopyTableRowStyleCopy", ws);

                    Assert.AreEqual(ExcelBorderStyle.Dashed, tbl.HeaderRowStyle.Border.Bottom.Style);
                    Assert.AreEqual(Color.Black.ToArgb(), tbl.HeaderRowStyle.Border.Bottom.Color.Color.Value.ToArgb());
                    Assert.AreEqual(Color.Red.ToArgb(), tbl.DataStyle.Font.Color.Color.Value.ToArgb());
                    Assert.AreEqual(Color.Green.ToArgb(), tbl.Columns[0].DataStyle.Font.Color.Color.Value.ToArgb());
                    SaveWorkbook("TableDxfCopy.xlsx", p2);
                }
            }
        }
    }
}


