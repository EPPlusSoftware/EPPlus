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
using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using OfficeOpenXml.Table.PivotTable;
using System.Text;
using System.Globalization;
using OfficeOpenXml.Drawing;
using System.Threading;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using Newtonsoft.Json;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Style.Coloring;

namespace EPPlusTest
{
    /// <summary>
    /// This class contains testcases for issues on Codeplex and Github.
    /// All tests requiering an template should be set to ignored as it's not practical to include all xlsx templates in the project.
    /// </summary>
    [TestClass]
    public class Issues : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestInitialize]
        public void Initialize()
        {
        }
        [TestMethod]
        public void Issue15041()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = 202100083;
                ws.Cells["A1"].Style.Numberformat.Format = "00\\.00\\.00\\.000\\.0";
                Assert.AreEqual("02.02.10.008.3", ws.Cells["A1"].Text);
                ws.Dispose();
            }
        }
        [TestMethod]
        public void Issue15031()
        {
            var d = OfficeOpenXml.Utils.ConvertUtil.GetValueDouble(new TimeSpan(35, 59, 1));
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = d;
                ws.Cells["A1"].Style.Numberformat.Format = "[t]:mm:ss";
                ws.Dispose();
            }
        }
        [TestMethod]
        public void Issue15022()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells.AutoFitColumns();
                ws.Cells["A1"].Style.Numberformat.Format = "0";
                ws.Cells.AutoFitColumns();
            }
        }
        [TestMethod]
        public void Issue15056()
        {
            using (var ep = OpenPackage(@"output.xlsx", true))
            {
                var s = ep.Workbook.Worksheets.Add("test");
                s.Cells["A1:A2"].Formula = ""; // or null, or non-empty whitespace, with same result
                ep.Save();
            }
        }
        [TestMethod]
        public void Issue15113()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1"].Value = " Performance Update";
            ws.Cells["A1:H1"].Merge = true;
            ws.Cells["A1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            ws.Cells["A1:H1"].Style.Font.Size = 14;
            ws.Cells["A1:H1"].Style.Font.Color.SetColor(Color.Red);
            ws.Cells["A1:H1"].Style.Font.Bold = true;
            SaveWorkbook(@"merge.xlsx", p);
            p.Dispose();
        }
        [TestMethod]
        public void Issue15141()
        {
            using (ExcelPackage package = new ExcelPackage())
            using (ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Test"))
            {
                sheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                sheet.Cells[1, 1, 1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[1, 5, 2, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                ExcelColumn column = sheet.Column(3); // fails with exception
            }
        }

        [TestMethod]
        public void Issue15123()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            using (var dt = new DataTable())
            {
                dt.Columns.Add("String", typeof(string));
                dt.Columns.Add("Int", typeof(int));
                dt.Columns.Add("Bool", typeof(bool));
                dt.Columns.Add("Double", typeof(double));
                dt.Columns.Add("Date", typeof(DateTime));

                var dr = dt.NewRow();
                dr[0] = "Row1";
                dr[1] = 1;
                dr[2] = true;
                dr[3] = 1.5;
                dr[4] = new DateTime(2014, 12, 30);
                dt.Rows.Add(dr);

                dr = dt.NewRow();
                dr[0] = "Row2";
                dr[1] = 2;
                dr[2] = false;
                dr[3] = 2.25;
                dr[4] = new DateTime(2014, 12, 31);
                dt.Rows.Add(dr);

                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.Cells["D2:D3"].Style.Numberformat.Format = "(* #,##0.00);_(* (#,##0.00);_(* \"-\"??_);(@)";

                ws.Cells["E2:E3"].Style.Numberformat.Format = "mm/dd/yyyy";
                ws.Cells.AutoFitColumns();
                Assert.AreNotEqual(ws.Cells[2, 5].Text, "");
            }
        }
        [TestMethod]
        public void Issue15128()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1"].Value = 1;
            ws.Cells["B1"].Value = 2;
            ws.Cells["B2"].Formula = "A1+$B$1";
            ws.Cells["C1"].Value = "Test";
            ws.Cells["A1:B2"].Copy(ws.Cells["C1"]);
            ws.Cells["B2"].Copy(ws.Cells["D1"]);
            SaveWorkbook("Copy.xlsx", p);
            p.Dispose();
        }

        [TestMethod]
        public void IssueMergedCells()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1:A5,C1:C8"].Merge = true;
            ws.Cells["C1:C8"].Merge = false;
            ws.Cells["A1:A8"].Merge = false;
            p.Dispose();
        }

        public class cls1
        {
            public int prop1 { get; set; }
        }

        public class cls2 : cls1
        {
            public string prop2 { get; set; }
        }
        [TestMethod]
        public void LoadFromColIssue()
        {
            var l = new List<cls1>();

            l.Add(new cls1() { prop1 = 1 });
            l.Add(new cls2() { prop1 = 1, prop2 = "test1" });

            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Test");

            ws.Cells["A1"].LoadFromCollection(l, true, TableStyles.Light16, BindingFlags.Instance | BindingFlags.Public,
                new MemberInfo[] { typeof(cls2).GetProperty("prop2") });
        }

        [TestMethod]
        public void Issue15168()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Test");
                ws.Cells[1, 1].Value = "A1";
                ws.Cells[2, 1].Value = "A2";

                ws.Cells[2, 1].Value = ws.Cells[1, 1].Value;
                Assert.AreEqual("A1", ws.Cells[1, 1].Value);
            }
        }
        [TestMethod]
        public void Issue15179()
        {
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("MergeDeleteBug");
                ws.Cells["E3:F3"].Merge = true;
                ws.Cells["E3:F3"].Merge = false;
                ws.DeleteRow(2, 6);
                ws.Cells["A1"].Value = 0;
                var s = ws.Cells["A1"].Value.ToString();

            }
        }
        [TestMethod]
        public void Issue15212()
        {
            var s = "_(\"R$ \"* #,##0.00_);_(\"R$ \"* (#,##0.00);_(\"R$ \"* \"-\"??_);_(@_) )";
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("StyleBug");
                ws.Cells["A1"].Value = 5698633.64;
                ws.Cells["A1"].Style.Numberformat.Format = s;
                var t = ws.Cells["A1"].Text;
            }
        }

        [TestMethod]
        /**** Pivottable issue ****/
        public void Issue()
        {

            using (var p = OpenPackage("pivottable.xlsx", true))
            {
                LoadData(p);
                BuildPivotTable1(p);
                BuildPivotTable2(p);
                p.Save();
            }
        }

        private void LoadData(ExcelPackage p)
        {
            // add a new worksheet to the empty workbook
            ExcelWorksheet wsData = p.Workbook.Worksheets.Add("Data");
            //Add the headers
            wsData.Cells[1, 1].Value = "INVOICE_DATE";
            wsData.Cells[1, 2].Value = "TOTAL_INVOICE_PRICE";
            wsData.Cells[1, 3].Value = "EXTENDED_PRICE_VARIANCE";
            wsData.Cells[1, 4].Value = "AUDIT_LINE_STATUS";
            wsData.Cells[1, 5].Value = "RESOLUTION_STATUS";
            wsData.Cells[1, 6].Value = "COUNT";

            //Add some items...
            wsData.Cells["A2"].Value = Convert.ToDateTime("04/2/2012");
            wsData.Cells["B2"].Value = 33.63;
            wsData.Cells["C2"].Value = (-.87);
            wsData.Cells["D2"].Value = "Unfavorable Price Variance";
            wsData.Cells["E2"].Value = "Pending";
            wsData.Cells["F2"].Value = 1;

            wsData.Cells["A3"].Value = Convert.ToDateTime("04/2/2012");
            wsData.Cells["B3"].Value = 43.14;
            wsData.Cells["C3"].Value = (-1.29);
            wsData.Cells["D3"].Value = "Unfavorable Price Variance";
            wsData.Cells["E3"].Value = "Pending";
            wsData.Cells["F3"].Value = 1;

            wsData.Cells["A4"].Value = Convert.ToDateTime("11/8/2011");
            wsData.Cells["B4"].Value = 55;
            wsData.Cells["C4"].Value = (-2.87);
            wsData.Cells["D4"].Value = "Unfavorable Price Variance";
            wsData.Cells["E4"].Value = "Pending";
            wsData.Cells["F4"].Value = 1;

            wsData.Cells["A5"].Value = Convert.ToDateTime("11/8/2011");
            wsData.Cells["B5"].Value = 38.72;
            wsData.Cells["C5"].Value = (-5.00);
            wsData.Cells["D5"].Value = "Unfavorable Price Variance";
            wsData.Cells["E5"].Value = "Pending";
            wsData.Cells["F5"].Value = 1;

            wsData.Cells["A6"].Value = Convert.ToDateTime("3/4/2011");
            wsData.Cells["B6"].Value = 77.44;
            wsData.Cells["C6"].Value = (-1.55);
            wsData.Cells["D6"].Value = "Unfavorable Price Variance";
            wsData.Cells["E6"].Value = "Pending";
            wsData.Cells["F6"].Value = 1;

            wsData.Cells["A7"].Value = Convert.ToDateTime("3/4/2011");
            wsData.Cells["B7"].Value = 127.55;
            wsData.Cells["C7"].Value = (-10.50);
            wsData.Cells["D7"].Value = "Unfavorable Price Variance";
            wsData.Cells["E7"].Value = "Pending";
            wsData.Cells["F7"].Value = 1;

            using (var range = wsData.Cells[2, 1, 7, 1])
            {
                range.Style.Numberformat.Format = "mm-dd-yy";
            }

            wsData.Cells.AutoFitColumns(0);
        }
        private void BuildPivotTable1(ExcelPackage p)
        {
            var wsData = p.Workbook.Worksheets["Data"];
            var totalRows = wsData.Dimension.Address;
            ExcelRange data = wsData.Cells[totalRows];

            var wsAuditPivot = p.Workbook.Worksheets.Add("Pivot1");

            var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit1");
            pivotTable1.ColumnGrandTotals = true;
            var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


            rowField.AddDateGrouping(eDateGroupBy.Years);
            var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
            yearField.Name = "Year";

            var rowField2 = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

            var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
            TotalSpend.Name = "Total Spend";
            TotalSpend.Format = "$##,##0";


            var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
            CountInvoicePrice.Name = "Total Lines";
            CountInvoicePrice.Format = "##,##0";

            pivotTable1.DataOnRows = false;
        }

        private void BuildPivotTable2(ExcelPackage p)
        {
            var wsData = p.Workbook.Worksheets["Data"];
            var totalRows = wsData.Dimension.Address;
            ExcelRange data = wsData.Cells[totalRows];

            var wsAuditPivot = p.Workbook.Worksheets.Add("Pivot2");

            var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit2");
            pivotTable1.ColumnGrandTotals = true;
            var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


            rowField.AddDateGrouping(eDateGroupBy.Years);
            var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
            yearField.Name = "Year";

            var rowField2 = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

            var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
            TotalSpend.Name = "Total Spend";
            TotalSpend.Format = "$##,##0";


            var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
            CountInvoicePrice.Name = "Total Lines";
            CountInvoicePrice.Format = "##,##0";

            pivotTable1.DataOnRows = false;
        }

        [TestMethod]
        public void Issue15377()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws1");
                ws.Cells["A1"].Value = (double?)1;
                var v = ws.GetValue<double?>(1, 1);
            }
        }
        [TestMethod]
        public void Issue15374()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RT");
                var r = ws.Cells["A1"];
                r.RichText.Text = "Cell 1";
                r["A2"].RichText.Add("Cell 2");
                SaveWorkbook(@"rt.xlsx", p);
            }
        }
        [TestMethod]
        public void IssueTranslate()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Trans");
                ws.Cells["A1:A2"].Formula = "IF(1=1, \"A's B C\",\"D\") ";
                var fr = ws.Cells["A1:A2"].FormulaR1C1;
                ws.Cells["A1:A2"].FormulaR1C1 = fr;
                Assert.AreEqual("IF(1=1,\"A's B C\",\"D\")", ws.Cells["A2"].Formula);
            }
        }
        [TestMethod]
        public void Issue15397()
        {
            using (var p = new ExcelPackage())
            {
                var workSheet = p.Workbook.Worksheets.Add("styleerror");
                workSheet.Cells["F:G"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["F:G"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                workSheet.Cells["A:A,C:C"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["A:A,C:C"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                //And then: 

                workSheet.Cells["A:H"].Style.Font.Color.SetColor(Color.Blue);

                workSheet.Cells["I:I"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["I:I"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                workSheet.Cells["I2"].Style.Fill.BackgroundColor.SetColor(Color.Green);
                workSheet.Cells["I4"].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                workSheet.Cells["I9"].Style.Fill.BackgroundColor.SetColor(Color.Pink);

                workSheet.InsertColumn(2, 2, 9);
                workSheet.Column(45).Width = 0;

                SaveWorkbook(@"styleerror.xlsx", p);
            }
        }
        [TestMethod]
        public void Issuer14801()
        {
            using (var p = new ExcelPackage())
            {
                var workSheet = p.Workbook.Worksheets.Add("rterror");
                var cell = workSheet.Cells["A1"];
                cell.RichText.Add("toto: ");
                cell.RichText[0].PreserveSpace = true;
                cell.RichText[0].Bold = true;
                cell.RichText.Add("tata");
                cell.RichText[1].Bold = false;
                cell.RichText[1].Color = Color.Green;
                SaveWorkbook(@"rtpreserve.xlsx", p);
            }
        }
        [TestMethod]
        public void Issuer15445()
        {
            using (var p = new ExcelPackage())
            {
                var ws1 = p.Workbook.Worksheets.Add("ws1");
                var ws2 = p.Workbook.Worksheets.Add("ws2");
                ws2.View.SelectedRange = "A1:B3 D12:D15";
                ws2.View.ActiveCell = "D15";
                SaveWorkbook(@"activeCell.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue15438()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Test");
                var c = ws.Cells["A1"].Style.Font.Color;
                c.Indexed = 3;
                Assert.AreEqual(c.LookupColor(c), "#FF00FF00");
            }
        }
        public static byte[] ReadTemplateFile(string templateName)
        {
            byte[] templateFIle;
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                using (var sw = new System.IO.FileStream(templateName, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                {
                    byte[] buffer = new byte[2048];
                    int bytesRead;
                    while ((bytesRead = sw.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ms.Write(buffer, 0, bytesRead);
                    }
                }
                ms.Position = 0;
                templateFIle = ms.ToArray();
            }
            return templateFIle;

        }

        [TestMethod]
        public void Issue15455()
        {
            using (var pck = new ExcelPackage())
            {

                var sheet1 = pck.Workbook.Worksheets.Add("sheet1");
                var sheet2 = pck.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells["C2"].Value = 3;
                sheet1.Cells["C3"].Formula = "VLOOKUP(E1, Sheet2!A1:D6, C2, 0)";
                sheet1.Cells["E1"].Value = "d";

                sheet2.Cells["A1"].Value = "d";
                sheet2.Cells["C1"].Value = "dg";
                pck.Workbook.Calculate();
                var c3 = sheet1.Cells["C3"].Value;
                Assert.AreEqual("dg", c3);
            }
        }

        [TestMethod]
        public void Issue15548_SumIfsShouldHandleGaps()
        {
            using (var package = new ExcelPackage())
            {
                var test = package.Workbook.Worksheets.Add("Test");

                test.Cells["A1"].Value = 1;
                test.Cells["B1"].Value = "A";

                //test.Cells["A2"] is default
                test.Cells["B2"].Value = "A";

                test.Cells["A3"].Value = 1;
                test.Cells["B4"].Value = "B";

                test.Cells["D2"].Formula = "SUMIFS(A1:A3,B1:B3,\"A\")";

                test.Calculate();

                var result = test.Cells["D2"].GetValue<int>();

                Assert.AreEqual(1, result, string.Format("Expected 1, got {0}", result));
            }
        }
        [TestMethod]
        public void Issue15548_SumIfsShouldHandleBadData()
        {
            using (var package = new ExcelPackage())
            {
                var test = package.Workbook.Worksheets.Add("Test");

                test.Cells["A1"].Value = 1;
                test.Cells["B1"].Value = "A";

                test.Cells["A2"].Value = "Not a number";
                test.Cells["B2"].Value = "A";

                test.Cells["A3"].Value = 1;
                test.Cells["B4"].Value = "B";

                test.Cells["D2"].Formula = "SUMIFS(A1:A3,B1:B3,\"A\")";

                test.Calculate();

                var result = test.Cells["D2"].GetValue<int>();

                Assert.AreEqual(1, result, string.Format("Expected 1, got {0}", result));
            }
        }
        [TestMethod]
        public void Issue63() // See https://github.com/JanKallman/EPPlus/issues/63
        {
            using (var p1 = new ExcelPackage())
            {
                ExcelWorksheet ws = p1.Workbook.Worksheets.Add("ArrayTest");
                ws.Cells["A1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Cells["A3"].Value = 3;
                ws.Cells["B1:B3"].CreateArrayFormula("A1:A3");
                p1.Save();

                // Test: basic support to recognize array formulas after reading Excel workbook file
                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    Assert.AreEqual("A1:A3", p1.Workbook.Worksheets["ArrayTest"].Cells["B1"].Formula);
                    Assert.IsTrue(p1.Workbook.Worksheets["ArrayTest"].Cells["B1"].IsArrayFormula);
                }
            }
        }
        [TestMethod]
        public void Issue61()
        {
            DataTable table1 = new DataTable("TestTable");
            table1.Columns.Add("name");
            table1.Columns.Add("id");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("i61");
                ws.Cells["A1"].LoadFromDataTable(table1, true);
            }

        }
        [TestMethod]
        public void Issue57()
        {
            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("test");
            ws.Cells["A1"].LoadFromArrays(Enumerable.Empty<object[]>());
        }
        [TestMethod]
        public void Issue66()
        {

            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("Test!");
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Formula = "A1";
                var wb = pck.Workbook;
                wb.Names.Add("Name1", ws.Cells["A1:A2"]);
                ws.Names.Add("Name2", ws.Cells["A1"]);
                pck.Save();
                using (var pck2 = new ExcelPackage(pck.Stream))
                {
                    ws = pck2.Workbook.Worksheets["Test!"];

                }
            }
        }
        /// <summary>
        /// Creating a new ExcelPackage with an external stream should not dispose of 
        /// that external stream. That is the responsibility of the caller.
        /// Note: This test would pass with EPPlus 4.1.1. In 4.5.1 the line CloseStream() was added
        /// to the ExcelPackage.Dispose() method. That line is redundant with the line before, 
        /// _stream.Close() except that _stream.Close() is only called if the _stream is NOT
        /// an External Stream (and several other conditions).
        /// Note that CloseStream() doesn't do anything different than _stream.Close().
        /// </summary>
        [TestMethod]
        public void Issue184_Disposing_External_Stream()
        {
            // Arrange
            var stream = new MemoryStream();

            using (var excelPackage = new ExcelPackage(stream))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Issue 184");
                worksheet.Cells[1, 1].Value = "Hello EPPlus!";
                excelPackage.SaveAs(stream);
                // Act
            } // This dispose should not dispose of stream.

            // Assert
            Assert.IsTrue(stream.Length > 0);
        }
        [TestMethod]
        public void Issue204()
        {
            using (var pack = new ExcelPackage())
            {
                //create sheets
                var sheet1 = pack.Workbook.Worksheets.Add("Sheet 1");
                var sheet2 = pack.Workbook.Worksheets.Add("Sheet 2");
                //set some default values
                sheet1.Cells[1, 1].Value = 1;
                sheet2.Cells[1, 1].Value = 2;
                //fill the formula
                var formula = string.Format("'{0}'!R1C1", sheet1.Name);

                var cell = sheet2.Cells[2, 1];
                cell.FormulaR1C1 = formula;
                //Formula should remain the same
                Assert.AreEqual(formula.ToUpper(), cell.FormulaR1C1.ToUpper());
            }
        }
        [TestMethod, Ignore]
        public void Issue170()
        {
            using (var p = OpenTemplatePackage("print_titles_170.xlsx"))
            {
                p.Compatibility.IsWorksheets1Based = false;
                ExcelWorksheet sheet = p.Workbook.Worksheets[0];

                sheet.PrinterSettings.RepeatColumns = new ExcelAddress("$A:$C");
                sheet.PrinterSettings.RepeatRows = new ExcelAddress("$1:$3");

                SaveWorkbook("print_titles_170-Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue172()
        {
            var pck = OpenTemplatePackage("quest.xlsx");
            foreach (var ws in pck.Workbook.Worksheets)
            {
                Console.WriteLine(ws.Name);
            }

            pck.Dispose();
        }

        [TestMethod]
        public void Issue219()
        {
            using (var p = OpenTemplatePackage("issueFile.xlsx"))
            {
                foreach (var ws in p.Workbook.Worksheets)
                {
                    Console.WriteLine(ws.Name);
                }
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void Issue234()
        {
            using (var s = new MemoryStream())
            {
                var data = Encoding.UTF8.GetBytes("Bad data").ToArray();
                s.Write(data, 0, data.Length);
                var package = new ExcelPackage(s);
            }
        }

        [TestMethod]
        public void WorksheetNameWithSingeQuote()
        {
            var pck = OpenPackage("sheetname_pbl.xlsx", true);
            var ws = pck.Workbook.Worksheets.Add("Deal's History");
            var a = ws.Cells["A:B"];
            ws.AutoFilterAddress = ws.Cells["A1:C3"];
            pck.Workbook.Names.Add("Test", ws.Cells["B1:D2"]);
            var name = a.WorkSheetName;

            var a2 = new ExcelAddress("'Deal''s History'!a1:a3");
            Assert.AreEqual(a2.WorkSheetName, "Deal's History");
            pck.Save();
            pck.Dispose();

        }
        [ExpectedException(typeof(ArgumentException))]
        [TestMethod]
        public void Issue233()
        {
            //get some test data
            var cars = Car.GenerateList();

            var pck = OpenPackage("issue233.xlsx", true);

            var sheetName = "Summary_GLEDHOWSUGARCO![]()PTY";

            //Create the worksheet 
            var sheet = pck.Workbook.Worksheets.Add(sheetName);

            //Read the data into a range
            var range = sheet.Cells["A1"].LoadFromCollection(cars, true);

            //Make the range a table
            var tbl = sheet.Tables.Add(range, $"data{sheetName}");
            tbl.ShowTotal = true;
            tbl.Columns["ReleaseYear"].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;

            //save and dispose
            pck.Save();
            pck.Dispose();
        }
        public class Car
        {
            public int Id { get; set; }
            public string Make { get; set; }
            public string Model { get; set; }
            public int ReleaseYear { get; set; }

            public Car(int id, string make, string model, int releaseYear)
            {
                Id = id;
                Make = make;
                Model = model;
                ReleaseYear = releaseYear;
            }

            internal static List<Car> GenerateList()
            {
                return new List<Car>
            {
				//random data
				new Car(1,"Toyota", "Carolla", 1950),
                new Car(2,"Toyota", "Yaris", 2000),
                new Car(3,"Toyota", "Hilux", 1990),
                new Car(4,"Nissan", "Juke", 2010),
                new Car(5,"Nissan", "Trail Blazer", 1995),
                new Car(6,"Nissan", "Micra", 2018),
                new Car(7,"BMW", "M3", 1980),
                new Car(8,"BMW", "X5", 2008),
                new Car(9,"BMW", "M6", 2003),
                new Car(10,"Merc", "S Class", 2001)
            };
            }
        }
        [TestMethod]
        public void Issue236()
        {
            using (var p = OpenTemplatePackage("Issue236.xlsx"))
            {
                p.Workbook.Worksheets["Sheet1"].Cells[7, 10].AddComment("test", "Author");
                SaveWorkbook("Issue236-Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue228()
        {
            using (var p = OpenTemplatePackage("Font55.xlsx"))
            {
                var ws = p.Workbook.Worksheets["Sheet1"];
                var d = ws.Drawings.AddShape("Shape1", eShapeStyle.Diamond);
                ws.Cells["A1"].Value = "tasetraser";
                ws.Cells.AutoFitColumns();
                SaveWorkbook("Font55-Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue241()
        {
            var pck = OpenPackage("issue241", true);
            var wks = pck.Workbook.Worksheets.Add("test");
            wks.DefaultRowHeight = 35;
            pck.Save();
            pck.Dispose();
        }
        [TestMethod]
        public void Issue195()
        {
            using (var pkg = new OfficeOpenXml.ExcelPackage())
            {
                var sheet = pkg.Workbook.Worksheets.Add("Sheet1");
                var defaultStyle = pkg.Workbook.Styles.CreateNamedStyle("Default");
                defaultStyle.Style.Font.Name = "Arial";
                defaultStyle.Style.Font.Size = 18;
                defaultStyle.Style.Font.UnderLine = true;
                var boldStyle = pkg.Workbook.Styles.CreateNamedStyle("Bold", defaultStyle.Style);
                boldStyle.Style.Font.Color.SetColor(Color.Red);

                Assert.AreEqual("Arial", defaultStyle.Style.Font.Name);
                Assert.AreEqual(18, defaultStyle.Style.Font.Size);

                Assert.AreEqual("Arial", boldStyle.Style.Font.Name);
                Assert.AreEqual(18, boldStyle.Style.Font.Size);
                Assert.AreEqual(boldStyle.Style.Font.Color.Rgb, "FFFF0000");

                SaveWorkbook("DefaultStyle.xlsx", pkg);
            }
        }
        [TestMethod]
        public void Issue332()
        {
            InitBase();
            var pkg = OpenPackage("Hyperlink.xlsx", true);
            var ws = pkg.Workbook.Worksheets.Add("Hyperlink");
            ws.Cells["A1"].Hyperlink = new ExcelHyperLink("A2", "A2");
            pkg.Save();
        }
        [TestMethod]
        public void Issue332_2()
        {
            InitBase();
            var pkg = OpenPackage("Hyperlink.xlsx");
            var ws = pkg.Workbook.Worksheets["Hyperlink"];
            Assert.IsNotNull(ws.Cells["A1"].Hyperlink);
        }
        [TestMethod]
        public void Issuer246()
        {
            InitBase();
            var pkg = OpenPackage("issue246.xlsx", true);
            var ws = pkg.Workbook.Worksheets.Add("DateFormat");
            ws.Cells["A1"].Value = 43465;
            ws.Cells["A1"].Style.Numberformat.Format = @"[$-F800]dddd,\ mmmm\ dd,\ yyyy";
            pkg.Save();

            pkg = OpenPackage("issue246.xlsx");
            ws = pkg.Workbook.Worksheets["DateFormat"];
            var pCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-Se");
            Assert.AreEqual(ws.Cells["A1"].Text, "den 31 december 2018");
            Assert.AreEqual(ws.GetValue<DateTime>(1, 1), new DateTime(2018, 12, 31));
            System.Threading.Thread.CurrentThread.CurrentCulture = pCulture;
        }
        [TestMethod]
        public void Issue347()
        {
            var package = OpenTemplatePackage("Issue327.xlsx");
            var templateWS = package.Workbook.Worksheets["Template"];
            //package.Workbook.Worksheets.Add("NewWs", templateWS);
            package.Workbook.Worksheets.Delete(templateWS);
        }
        [TestMethod]
        public void Issue348()
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("S1");
                string formula = "VLOOKUP(C2,A:B,1,0)";
                ws.Cells[2, 4].Formula = formula;
                var t1 = ws.Cells[2, 4].FormulaR1C1; // VLOOKUP(C2,C[-3]:C[-2],1,0)
                ws.Cells[2, 5].FormulaR1C1 = ws.Cells[2, 4].FormulaR1C1;
                var t2 = ws.Cells[2, 5].FormulaR1C1; // VLOOKUP(C2,C[-3]**:B:C:C**,1,0)   //unexpected value here
            }
        }

        [TestMethod]
        public void Issue367()
        {
            using (var pck = OpenTemplatePackage(@"ProductFunctionTest.xlsx"))
            {
                var sheet = pck.Workbook.Worksheets.First();
                //sheet.Cells["B13"].Value = null;
                sheet.Cells["B14"].Value = 11;
                sheet.Cells["B15"].Value = 13;
                sheet.Cells["B16"].Formula = "Product(B13:B15)";
                sheet.Calculate();

                Assert.AreEqual(0d, sheet.Cells["B16"].Value);
            }
        }
        [TestMethod]
        public void Issue345()
        {
            using (ExcelPackage package = OpenTemplatePackage("issue345.xlsx"))
            {
                var worksheet = package.Workbook.Worksheets["test"];
                int[] sortColumns = new int[1];
                sortColumns[0] = 0;
                worksheet.Cells["A2:A30864"].Sort(sortColumns);
                package.Save();
            }
        }
        [TestMethod]
        public void Issue387()
        {

            using (ExcelPackage package = OpenTemplatePackage("issue345.xlsx"))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets.Add("One");

                worksheet.Cells[1, 3].Value = "Hello";
                var cells = worksheet.Cells["A3"];

                worksheet.Names.Add("R0", cells);
                workbook.Names.Add("Q0", cells);
            }
        }
        [TestMethod]
        public void Issue333()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-SE");
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("TextBug");
                ws.Cells["A1"].Value = new DateTime(2019, 3, 7);
                ws.Cells["A1"].Style.Numberformat.Format = "mm-dd-yy";

                Assert.AreEqual("2019-03-07", ws.Cells["A1"].Text);
            }
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("TextBug");
                ws.Cells["A1"].Value = new DateTime(2019, 3, 7);
                ws.Cells["A1"].Style.Numberformat.Format = "mm-dd-yy";

                Assert.AreEqual("3/7/2019", ws.Cells["A1"].Text);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }
        [TestMethod]
        public void Issue445()
        {
            ExcelPackage p = new ExcelPackage();
            ExcelWorksheet ws = p.Workbook.Worksheets.Add("AutoFit"); //<-- This line takes forever. The process hangs.
            ws.Cells[1, 1].Value = new string('a', 50000);
            ws.Cells[1, 1].AutoFitColumns();
        }
        [TestMethod]
        public void Issue551()
        {
            using (var p = OpenTemplatePackage("Submittal.Extract.5.ton.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveWorkbook("Submittal.Extract.5.ton_Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue558()
        {
            using (var p = OpenTemplatePackage("GoogleSpreadsheet.xlsx"))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets[0];
                p.Workbook.Worksheets.Copy(ws.Name, "NewName");
                SaveWorkbook("GoogleSpreadsheet-Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue520()
        {
            using (var p = OpenTemplatePackage("template_slim.xlsx"))
            {

                var workSheet = p.Workbook.Worksheets[0];
                workSheet.Cells["B5"].LoadFromArrays(new List<object[]> { new object[] { "xx", "Name", 1, 2, 3, 5, 6, 7 } });

                SaveWorkbook("ErrorStyle0.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue510()
        {
            using (var p = OpenTemplatePackage("Error.Opening.with.EPPLus.xlsx"))
            {

                var workSheet = p.Workbook.Worksheets[0];

                SaveWorkbook("Issue510.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue464()
        {
            using (var p1 = OpenTemplatePackage("Sample_Cond_Format.xlsx"))
            {
                var ws = p1.Workbook.Worksheets[0];
                using (var p2 = new ExcelPackage())
                {
                    var ws2 = p2.Workbook.Worksheets.Add("Test", ws);
                    foreach (var cf in ws2.ConditionalFormatting)
                    {

                    }
                    SaveWorkbook("CondCopy.xlsx", p2);
                }
            }
        }
        [TestMethod]
        public void Issue436()
        {
            using (var p1 = OpenTemplatePackage("issue436.xlsx"))
            {
                var ws = p1.Workbook.Worksheets[0];
                Assert.IsNotNull(((ExcelShape)ws.Drawings[0]).Text);
            }
        }
        [TestMethod]
        public void Issue425()
        {
            using (var p1 = OpenTemplatePackage("issue425.xlsm"))
            {
                var ws = p1.Workbook.Worksheets[1];

                p1.Workbook.Worksheets.Add("NewNotCopied");
                p1.Workbook.Worksheets.Add("NewCopied", ws);

                SaveWorkbook("issue425.xlsm", p1);
            }
        }
        [TestMethod]
        public void Issue422()
        {
            using (var p1 = OpenTemplatePackage("CustomFormula.xlsx"))
            {
                SaveWorkbook("issue422.xlsx", p1);
            }
        }

        [TestMethod]
        public void Issue625()
        {
            using (var p = OpenTemplatePackage("multiple_print_areas.xlsx"))
            {

                var workSheet = p.Workbook.Worksheets[0];

                SaveWorkbook("Issue625.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue403()
        {
            using (var p = OpenTemplatePackage("issue403.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveWorkbook("Issue403.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue39()
        {
            using (var p = OpenTemplatePackage("MyExcel.xlsx"))
            {
                var workSheet = p.Workbook.Worksheets[0];

                workSheet.InsertRow(8, 2, 8);
                SaveWorkbook("Issue39.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue70()
        {
            using (var p = OpenTemplatePackage("HiddenOO.xlsx"))
            {
                Assert.IsTrue(p.Workbook.Worksheets[0].Column(2).Hidden);
                SaveWorkbook("Issue70.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue72()
        {
            using (var p = OpenTemplatePackage("Issue72-Table.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual("COUNTIF(Base[Date],Calc[[#This Row],[Date]])", ws.Cells["F3"].Formula);
                SaveWorkbook("Issue72.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue54()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergeBug");

                var r = ws.Cells[1, 1, 1, 5];
                r.Merge = true;
                r.Value = "Header";
                SaveWorkbook("Issue54.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue55()
        {
            using (var p = new ExcelPackage())
            {
                var worksheet = p.Workbook.Worksheets.Add("DV test");
                var rangeToSet = ExcelCellBase.GetAddress(1, 3, ExcelPackage.MaxRows, 3);
                worksheet.Names.Add("ListName", worksheet.Cells["D1:D3"]);
                worksheet.Cells["D1"].Value = "A";
                worksheet.Cells["D2"].Value = "B";
                worksheet.Cells["D3"].Value = "C";
                var validation = worksheet.DataValidations.AddListValidation(rangeToSet);
                validation.Formula.ExcelFormula = $"=ListName";
                SaveWorkbook("dv.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue73()
        {
            using (var p = OpenTemplatePackage("Issue73.xlsx"))
            {
                var workSheet = p.Workbook.Worksheets[0];

                SaveWorkbook("Issue73Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue74()
        {
            using (var p = OpenTemplatePackage("Issue74.xlsx"))
            {
                var workSheet = p.Workbook.Worksheets[0];

                SaveWorkbook("Issue74Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue76()
        {
            using (var p = OpenTemplatePackage("Issue76.xlsx"))
            {
                var workSheet = p.Workbook.Worksheets[0];

                SaveWorkbook("Issue76Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue88()
        {
            using (var p = OpenTemplatePackage("Issue88.xlsm"))
            {
                var ws1 = p.Workbook.Worksheets[0];
                var ws2 = p.Workbook.Worksheets.Add("test", ws1);
                SaveWorkbook("Issue88Saved.xlsm", p);
            }
        }
        [TestMethod]
        public void Issue94()
        {
            using (var p = OpenTemplatePackage("Issue425.xlsm"))
            {
                p.Workbook.VbaProject.Remove();
                SaveWorkbook("Issue425.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue95()
        {
            using (var p = OpenTemplatePackage("Issue95.xlsx"))
            {
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue99()
        {
            using (var p = OpenTemplatePackage("Issue-99-2.xlsx"))
            {
                //var p2 = OpenPackage("Issue99-2Saved-new.xlsx", true);
                //var ws = p2.Workbook.Worksheets.Add("Picture");
                //ws.Drawings.AddPicture("Test1", Properties.Resources.Test1);
                //p.Workbook.Worksheets.Add("copy1", p.Workbook.Worksheets[0]);
                //p2.Workbook.Worksheets.Add("copy1", p.Workbook.Worksheets[0]);
                //p.Workbook.Worksheets.Add("copy2", p2.Workbook.Worksheets[0]);
                //SaveAndCleanup(p2);
                SaveWorkbook("Issue99-2Saved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue115()
        {
            using (var p = OpenPackage("Issue115.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("DefinedNamesIssue");
                p.Workbook.Names.Add("Name", ws.Cells["B6:D8,B10:D11"]);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue121()
        {
            using (var p = OpenTemplatePackage("Deployment aftaler.xlsx"))
            {
                var sheet = p.Workbook.Worksheets[0];
            }
        }

        [TestMethod, Ignore]
        public void SupportCase17()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\Issue17\BreakLinks3.xlsx")))
            {
                var stopwatch = Stopwatch.StartNew();
                p.Workbook.FormulaParserManager.AttachLogger(new FileInfo("c:\\temp\\formulalog.txt"));
                p.Workbook.Calculate();
                stopwatch.Stop();
                var ms = stopwatch.Elapsed.TotalSeconds;

            }
        }
        [TestMethod, Ignore]
        public void Issue17()
        {
            using (var p = OpenTemplatePackage("Excel Sample Circular Ref break links.xlsx"))
            {
                p.Workbook.Calculate();
            }
        }
        [TestMethod]
        public void Issue18()
        {
            using (var p = OpenTemplatePackage("000P020-SQ101_H0.xlsm"))
            {
                p.Workbook.Worksheets.Delete(0);
                p.Workbook.Worksheets.Delete(2);
                p.Workbook.Calculate();
                SaveWorkbook("null_issue_vba.xlsm", p);
            }
        }
        [TestMethod]
        public void Issue26()
        {
            using (var p = OpenTemplatePackage("Issue26.xlsx"))
            {
                SaveAndCleanup(p);
            }
            using (var p = OpenPackage("Issue26.xlsx"))
            {
                SaveWorkbook("Issue26-resaved.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue180()
        {
            var p1 = OpenTemplatePackage("Issue180-1.xlsm");
            var p2 = OpenTemplatePackage("Issue180-2.xlsm");
            p2.Workbook.Worksheets.Add(p1.Workbook.Worksheets[0].Name, p1.Workbook.Worksheets[0]);
            p2.SaveAs(new FileInfo("c:\\epplustest\\t.xlsm"));
        }
        [TestMethod]
        public void Issue34()
        {
            using (var p = OpenTemplatePackage("Issue34.xlsx"))
            {
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void Issue38()
        {
            using (var p = OpenTemplatePackage("pivottest.xlsx"))
            {
                Assert.AreEqual(1, p.Workbook.Worksheets[1].PivotTables.Count);
                var tbl = p.Workbook.Worksheets[0].Tables[0];
                var pt = p.Workbook.Worksheets[1].PivotTables[0];
                Assert.IsNotNull(p.Workbook.Worksheets[1].PivotTables[0].CacheDefinition);
                var s1 = pt.Fields[0].AddSlicer();
                s1.SetPosition(0, 500);
                var s2 = pt.Fields["OpenDate"].AddSlicer();
                pt.Fields["Distance"].Format = "#,##0.00";
                pt.Fields["Distance"].AddSlicer();
                s2.SetPosition(0, 500 + (int)s1._width);
                tbl.Columns["IsUser"].AddSlicer();
                pt.Fields["IsUser"].AddSlicer();

                SaveWorkbook("pivotTable2.xlsx", p);
            }
        }
        [TestMethod]
        public void Issue195_PivotTable()
        {
            using (var p = OpenTemplatePackage("Issue195.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue45()
        {
            using (var p = OpenPackage("LinkIssue.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1:A2"].Value = 1;
                ws.Cells["B1:B2"].Formula = $"VLOOKUP($A1,[externalBook.xlsx]Prices!$A:$H, 3, FALSE)";
                SaveAndCleanup(p);
            }

        }
        [TestMethod]
        public void EmfIssue()
        {
            using (var p = OpenTemplatePackage("emfIssue.xlsm"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue201()
        {
            using (var p = OpenTemplatePackage("book1.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual("0", ws.Cells["A1"].Text);
                Assert.AreEqual("-", ws.Cells["A2"].Text);
                Assert.AreEqual("0", ws.Cells["A3"].Text);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void IssueCellstore()
        {
            int START_ROW = 1;
            int CustomTemplateRowsOffset = 4;
            int rowCount = 34000;
            using (var package = OpenTemplatePackage("CellStoreIssue.xlsm"))
            {
                var worksheet = package.Workbook.Worksheets[0];
                worksheet.Cells["A5"].Value = "Test";
                worksheet.InsertRow(START_ROW + CustomTemplateRowsOffset, rowCount - 1, CustomTemplateRowsOffset + 1);
                Assert.AreEqual("Test", worksheet.Cells["A34004"].Value);
                //for (int k = START_ROW+CustomTemplateRowsOffset; k < rowCount; k++)
                //{
                //    worksheet.Cells[(START_ROW + CustomTemplateRowsOffset) + ":" + (START_ROW + CustomTemplateRowsOffset)]
                //        .Copy(worksheet.Cells[k + 1 + ":" + k + 1]);
                //}
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void Issue220()
        {
            using (var p = OpenTemplatePackage("Generated.with.EPPlus.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
            }
        }
        [TestMethod]
        public void Issue232()
        {
            using (var p = OpenTemplatePackage("pivotbug541.xlsx"))
            {
                var overviewSheet = p.Workbook.Worksheets["Overblik"];
                var serverSheet = p.Workbook.Worksheets["Servers"];
                var serverPivot = overviewSheet.PivotTables.Add(overviewSheet.Cells["A4"], serverSheet.Cells[serverSheet.Dimension.Address], "ServerPivot");
                p.Save();
            }
        }
        [TestMethod]
        public void Issue_234()
        {
            using (var p = OpenTemplatePackage("ExcelErrorFile.xlsx"))
            {
                var ws = p.Workbook.Worksheets["Leistung"];

                Assert.IsNull(ws.Cells["C65538"].Value);
                Assert.IsNull(ws.Cells["C71715"].Value);
                Assert.AreEqual(0D, ws.Cells["C71716"].Value);
                Assert.AreEqual(0D, ws.Cells["C71811"].Value);
                Assert.IsNull(ws.Cells["C71812"].Value);
                Assert.IsNull(ws.Cells["C77667"].Value);
                Assert.AreEqual(0D, ws.Cells["C77668"].Value);
            }
        }
        [TestMethod]
        public void InflateIssue()
        {
            using (var p = OpenPackage("inflateStart.xlsx", true))
            {
                var worksheet = p.Workbook.Worksheets.Add("Test");
                for (int i = 1; i <= 10; i++)
                {
                    worksheet.Cells[1, i].Hyperlink = new Uri("https://epplussoftware.com");
                    worksheet.Cells[1, i].Value = "Url " + worksheet.Cells[1, i].Address;
                }
                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    for (int i = 0; i < 10; i++)
                    {
                        p.Save();
                    }
                    SaveWorkbook("Inflate.xlsx", p2);
                }
            }
        }
        [TestMethod]
        public void DrawingSetFont()
        {
            using (var p = OpenPackage("DrawingSetFromFont.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Drawing1");
                var shape = ws.Drawings.AddShape("x", eShapeStyle.Rect);
                shape.Font.SetFromFont("Arial", 20);
                shape.Text = "Font";
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue_258()
        {
            using (var package = OpenTemplatePackage("Test.xlsx"))
            {
                var overviewSheet = package.Workbook.Worksheets["Overview"];
                if (overviewSheet != null)
                    package.Workbook.Worksheets.Delete(overviewSheet);
                overviewSheet = package.Workbook.Worksheets.Add("Overview");
                var serverSheet = package.Workbook.Worksheets["Servers"];
                var serverPivot = overviewSheet.PivotTables.Add(overviewSheet.Cells["A4"], serverSheet.Cells[serverSheet.Dimension.Address], "ServerPivot");
                var serverNameField = serverPivot.Fields["Name"];
                serverPivot.RowFields.Add(serverNameField);
                var standardBackupField = serverPivot.Fields["StandardBackup"];
                serverPivot.PageFields.Add(standardBackupField);
                standardBackupField.Items.Refresh();
                var items = standardBackupField.Items;
                items.SelectSingleItem(1); // <===== this one is to select only the "false" condition
                SaveWorkbook("Issue248.xlsx", package);
            }
        }
        [TestMethod]
        public void Issue_243()
        {
            using (var p = OpenPackage("formula.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("formula");
                ws.Cells["A1"].Value = "column1";
                ws.Cells["A2"].Value = 1;
                ws.Cells["A3"].Value = 2;
                ws.Cells["A4"].Value = 3;

                var tbl = ws.Tables.Add(ws.Cells["A1:A4"], "Table1");

                ws.Cells["B1"].Formula = "TEXTJOIN(\" | \", false, INDIRECT(\"Table1[#data]\"))";
                ws.Calculate();
                Assert.AreEqual("1 | 2 | 3", ws.Cells["B1"].Value);

                ws.Cells["B1"].Formula = "TEXTJOIN(\" | \", false, INDIRECT(\"Table1\"))";
                ws.Calculate();
                Assert.AreEqual("1 | 2 | 3", ws.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void IssueCommentInsert()
        {

            using (var p = OpenPackage("CommentInsert.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("CommentInsert");
                ws.Cells["A2"].AddComment("na", "test");
                Assert.AreEqual(1, ws.Comments.Count);

                ws.InsertRow(2, 1);
                ws.Cells["A3"].Insert(eShiftTypeInsert.Right);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue261()
        {
            using (var p = OpenTemplatePackage("issue261.xlsx"))
            {
                var ws = p.Workbook.Worksheets["data"];
                ws.Cells["A1"].Value = "test";
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue260()
        {
            using (var p = OpenTemplatePackage("issue260.xlsx"))
            {
                var workbook = p.Workbook;
                Console.WriteLine(workbook.Worksheets.Count);
            }
        }
        [TestMethod]
        public void Issue268()
        {
            using (var p = OpenPackage("Issue268.xlsx", true))
            {
                ExcelWorksheet formSheet = CreateFormSheet(p);
                var r1 = formSheet.Drawings.AddCheckBoxControl("OptionSingleRoom");
                r1.Text = "Single Room";
                r1.LinkedCell = formSheet.Cells["G7"];
                r1.SetPosition(5, 0, 1, 0);
                var tableSheet = p.Workbook.Worksheets.Add("Table");
                ExcelRange tableRange = formSheet.Cells[10, 20, 30, 22];
                ExcelTable faultsTable = formSheet.Tables.Add(tableRange, "FaultsTable");
                faultsTable.StyleName = "None";
                SaveAndCleanup(p);
            }
        }
        private static ExcelWorksheet CreateFormSheet(ExcelPackage package)
        {
            var formSheet = package.Workbook.Worksheets.Add("Form");
            formSheet.Cells["A1"].Value = "Room booking";
            formSheet.Cells["A1"].Style.Font.Size = 18;
            formSheet.Cells["A1"].Style.Font.Bold = true;
            return formSheet;
        }
        [TestMethod]
        public void Issue269()
        {
            var data = new List<TestDTO>();

            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("Sheet1");
                var r = sheet.Cells["A1"].LoadFromCollection(data, false);
                Assert.IsNull(r);
            }
        }
        [TestMethod]
        public void Issue272()
        {
            using (var p = OpenTemplatePackage("Issue272.xlsx"))
            {
                var workbook = p.Workbook;
                Console.WriteLine(workbook.Worksheets.Count);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void IssueS84()
        {
            using (var p = OpenTemplatePackage("XML in Cells.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var cell = ws.Cells["D43"];
                cell.Value += " ";

                ExcelRichText rtx = cell.RichText.Add("a");

                rtx.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;

                ws.Cells["D43:E44"].Value = new object[,] { { "Cell1", "Cell2" }, { "Cell21", "Cell22" } };

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void IssueS80()
        {
            using (var p = OpenTemplatePackage("Example - CANNOT OPEN EPPLUS.xlsx"))
            {
                var workbook = p.Workbook;
                SaveAndCleanup(p);
                new ExcelAddress("f");
            }
        }
        [TestMethod]
        public void IssueS91()
        {
            using (var p = OpenTemplatePackage("Tagging Template V14.xlsx"))
            {
                var ws = p.Workbook.Worksheets["Stacked Logs"];
                //Insert 2 rows extending the data validations. 
                ws.InsertRow(4, 2, 4);

                //Get the data validation of choice.
                var dv = ws.DataValidations[0].As.ListValidation;

                //Adjust the formula using the R1C1 translator...
                var formula = dv.Formula.ExcelFormula;
                var r1c1Formula = OfficeOpenXml.Core.R1C1Translator.ToR1C1Formula(formula, dv.Address.Start.Row, dv.Address.Start.Column);
                //Add one row to the formula
                var formulaRowPlus1 = OfficeOpenXml.Core.R1C1Translator.FromR1C1Formula(r1c1Formula, dv.Address.Start.Row + 1, dv.Address.Start.Column);

                SaveAndCleanup(p);
            }
        }
        public class Test
        {
            public int Value1 { get; set; }
            public int Value2 { get; set; }
            public int Value3 { get; set; }

        }
        [TestMethod]
        public void Issue284()
        {
            //1
            var report1 = new List<Test>
            {
                new Test{ Value1 = 1, Value2= 2, Value3=3 },
                new Test{ Value1 = 2, Value2= 3, Value3=4 },
                new Test{ Value1 = 5, Value2= 6, Value3=7 }
            };

            //3
            var report2 = new List<Test>
            {
                new Test{ Value1 = 0, Value2= 0, Value3=0 },
                new Test{ Value1 = 0, Value2= 0, Value3=0 },
                new Test{ Value1 = 0, Value2= 0, Value3=0 }
            };

            //4
            var report3 = new List<Test>
            {
                new Test{ Value1 = 3, Value2= 3, Value3=3 },
                new Test{ Value1 = 3, Value2= 3, Value3=3 },
                new Test{ Value1 = 3, Value2= 3, Value3=3 }
            };


            string workingDirectory = AppDomain.CurrentDomain.BaseDirectory;
            using (var excelFile = OpenTemplatePackage("issue284.xlsx"))
            {
                //Data1
                var worksheet = excelFile.Workbook.Worksheets["Test1"];
                ExcelRangeBase location = worksheet.Cells["A1"].LoadFromCollection(Collection: report1, PrintHeaders: true);
                var t = worksheet.Tables.Add(location, "mytestTbl");
                t.TableStyle = TableStyles.None;

                //Data2
                worksheet = excelFile.Workbook.Worksheets["Test2"];
                location = worksheet.Cells["A1"].LoadFromCollection(Collection: report2, PrintHeaders: true);
                worksheet.Tables.Add(location, "mytestsureTbl");

                //Data3
                location = worksheet.Cells["K1"].LoadFromCollection(Collection: report3, PrintHeaders: true);
                worksheet.Tables.Add(location, "Test3");

                var wsFirst = excelFile.Workbook.Worksheets["Test1"];

                wsFirst.Select();
                SaveAndCleanup(excelFile);
            }
        }
        [TestMethod]
        public void Ticket90()
        {
            using (var p = OpenTemplatePackage("Example - Calculate.xlsx"))
            {
                var sheet = p.Workbook.Worksheets["Others"];
                var fi = new FileInfo(@"c:\Temp\countiflog.txt");
                p.Workbook.FormulaParserManager.AttachLogger(fi);
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = OfficeOpenXml.FormulaParsing.PrecisionAndRoundingStrategy.Excel);
                p.Workbook.FormulaParserManager.DetachLogger();
                var result = sheet.Cells["R5"].Value;
                ExcelAddress a = new ExcelAddress();

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Ticket90_2()
        {
            using (var p = OpenTemplatePackage("s70.xlsx"))
            {
                p.Workbook.Calculate();
                Assert.AreEqual(7D, p.Workbook.Worksheets[0].Cells["P1"].Value);
                Assert.AreEqual(1D, p.Workbook.Worksheets[0].Cells["P2"].Value);
                Assert.AreEqual(0D, p.Workbook.Worksheets[0].Cells["P3"].Value);
            }
        }
        [TestMethod]
        public void Issue287()
        {
            using (var p = OpenTemplatePackage("issue287.xlsm"))
            {
                p.Workbook.CreateVBAProject();
                p.Save();
            }
        }
        [TestMethod]
        public void Issue309()
        {
            using (var p = OpenTemplatePackage("test1.xlsx"))
            {
                p.Save();
            }
        }
        [TestMethod]
        public void Issue274()
        {
            using (var p = OpenPackage("Issue274.xlsx", true))
            {
                var worksheet = p.Workbook.Worksheets.Add("PayrollData");

                //Add the headers
                worksheet.Cells[1, 1].Value = "Employee";
                worksheet.Cells[1, 2].Value = "HomeOffice";
                worksheet.Cells[1, 3].Value = "JobNo";
                worksheet.Cells[1, 4].Value = "Ordinary";
                worksheet.Cells[1, 5].Value = "TimeHalf";
                worksheet.Cells[1, 6].Value = "DoubleTime";
                worksheet.Cells[1, 7].Value = "ProductiveHrs";
                worksheet.Cells[1, 8].Value = "NonProductiveHrs";


                int cnt = 2;
                worksheet.Cells[cnt, 1].Value = "Steve";
                worksheet.Cells[cnt, 2].Value = "Binda";
                worksheet.Cells[cnt, 3].Value = "SW001";
                worksheet.Cells[cnt, 4].Value = 12.0;
                worksheet.Cells[cnt, 5].Value = 6.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 18.0;
                worksheet.Cells[cnt, 8].Value = 0.0;
                cnt++;
                worksheet.Cells[cnt, 1].Value = "Steve";
                worksheet.Cells[cnt, 2].Value = "Binda";
                worksheet.Cells[cnt, 3].Value = "SW002";
                worksheet.Cells[cnt, 4].Value = 7.0;
                worksheet.Cells[cnt, 5].Value = 0.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 7.0;
                worksheet.Cells[cnt, 8].Value = 0.0;
                cnt++;
                worksheet.Cells[cnt, 1].Value = "Steve";
                worksheet.Cells[cnt, 2].Value = "Binda";
                worksheet.Cells[cnt, 3].Value = "Admin";
                worksheet.Cells[cnt, 4].Value = 4.0;
                worksheet.Cells[cnt, 5].Value = 0.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 0.0;
                worksheet.Cells[cnt, 8].Value = 4.0;
                cnt++;
                worksheet.Cells[cnt, 1].Value = "Peter";
                worksheet.Cells[cnt, 2].Value = "Binda";
                worksheet.Cells[cnt, 3].Value = "SW001";
                worksheet.Cells[cnt, 4].Value = 12.0;
                worksheet.Cells[cnt, 5].Value = 6.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 18.0;
                worksheet.Cells[cnt, 8].Value = 0.0;
                cnt++;
                worksheet.Cells[cnt, 1].Value = "Peter";
                worksheet.Cells[cnt, 2].Value = "Binda";
                worksheet.Cells[cnt, 3].Value = "SW002";
                worksheet.Cells[cnt, 4].Value = 7.0;
                worksheet.Cells[cnt, 5].Value = 0.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 7.0;
                worksheet.Cells[cnt, 8].Value = 0.0;
                cnt++;
                worksheet.Cells[cnt, 1].Value = "Peter";
                worksheet.Cells[cnt, 2].Value = "Binda";
                worksheet.Cells[cnt, 3].Value = "Admin";
                worksheet.Cells[cnt, 4].Value = 4.0;
                worksheet.Cells[cnt, 5].Value = 0.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 0.0;
                worksheet.Cells[cnt, 8].Value = 4.0;
                cnt++;
                worksheet.Cells[cnt, 1].Value = "Brian";
                worksheet.Cells[cnt, 2].Value = "Sydney";
                worksheet.Cells[cnt, 3].Value = "SW001";
                worksheet.Cells[cnt, 4].Value = 12.0;
                worksheet.Cells[cnt, 5].Value = 6.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 18.0;
                worksheet.Cells[cnt, 8].Value = 0.0;
                cnt++;
                worksheet.Cells[cnt, 1].Value = "Brian";
                worksheet.Cells[cnt, 2].Value = "Binda";
                worksheet.Cells[cnt, 3].Value = "SW002";
                worksheet.Cells[cnt, 4].Value = 7.0;
                worksheet.Cells[cnt, 5].Value = 0.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 7.0;
                worksheet.Cells[cnt, 8].Value = 0.0;
                cnt++;
                worksheet.Cells[cnt, 1].Value = "Brian";
                worksheet.Cells[cnt, 2].Value = "Binda";
                worksheet.Cells[cnt, 3].Value = "Admin";
                worksheet.Cells[cnt, 4].Value = 4.0;
                worksheet.Cells[cnt, 5].Value = 0.0;
                worksheet.Cells[cnt, 6].Value = 0.0;
                worksheet.Cells[cnt, 7].Value = 0.0;
                worksheet.Cells[cnt, 8].Value = 4.0;

                cnt--;
                using (var range = worksheet.Cells[1, 1, 1, 8])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                    range.Style.Font.Color.SetColor(Color.White);
                }
                var dataRange = worksheet.Cells[1, 1, cnt, 8];
                ExcelTableCollection tblcollection = worksheet.Tables;
                ExcelTable table = tblcollection.Add(dataRange, "payrolldata");
                table.ShowHeader = true;
                table.ShowFilter = true;
                var wsPivot = p.Workbook.Worksheets.Add("Employee-Job");
                var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], dataRange, "ByEmployee");
                pivotTable.RowFields.Add(pivotTable.Fields["Employee"]);
                var rowField1 = pivotTable.RowFields.Add(pivotTable.Fields["HomeOffice"]);
                var rowField2 = pivotTable.RowFields.Add(pivotTable.Fields["JobNo"]);
                var calcField1 = pivotTable.Fields.AddCalculatedField("Productive", "'ProductiveHrs'/('ProductiveHrs'+'NonProductiveHrs')*100");
                calcField1.Format = "#,##0";
                ExcelPivotTableDataField dataField;
                dataField = pivotTable.DataFields.Add(pivotTable.Fields["Productive"]);
                dataField.Format = "#,##0.0";
                dataField.Name = "Productive2";

                dataField = pivotTable.DataFields.Add(pivotTable.Fields["Ordinary"]);
                dataField.Format = "#,##0.0";
                dataField = pivotTable.DataFields.Add(pivotTable.Fields["TimeHalf"]);
                dataField.Format = "#,##0.0";
                dataField = pivotTable.DataFields.Add(pivotTable.Fields["DoubleTime"]);
                dataField.Format = "#,##0.0";
                dataField = pivotTable.DataFields.Add(pivotTable.Fields["ProductiveHrs"]);
                dataField.Format = "#,##0.0";
                dataField = pivotTable.DataFields.Add(pivotTable.Fields["NonProductiveHrs"]);
                dataField.Format = "#,##0.0";
                pivotTable.DataOnRows = false;
                pivotTable.Compact = true;
                pivotTable.CompactData = true;
                pivotTable.OutlineData = true;
                //pivotTable.ShowDrill = true;
                //pivotTable.CacheDefinition.Refresh();
                pivotTable.Fields["Employee"].Items.ShowDetails(false);
                rowField1.Items.ShowDetails(false);
                worksheet.Cells.AutoFitColumns(0);

                // create macro's to collapse pivot table

                //p.Workbook.CreateVBAProject();
                //var sb = new StringBuilder();
                //sb.AppendLine("Private Sub Workbook_Open()");
                //sb.AppendLine("    Sheets(\"Employee-Job\").Select");
                //sb.AppendLine("    ActiveSheet.PivotTables(\"ByEmployee\").PivotFields(\"Employee\").ShowDetail = False");
                //sb.AppendLine("    ActiveSheet.PivotTables(\"ByEmployee\").PivotFields(\"HomeOffice\").ShowDetail = False");
                //sb.AppendLine("End Sub");
                //p.Workbook.CodeModule.Code = sb.ToString();
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void DeleteCommentIssue()
        {
            using (var p = OpenTemplatePackage("CommentDelete.xlsx"))
            {
                var ws = p.Workbook.Worksheets["S3"];
                ws.Comments[0].RichText.Add("T");
                ws.Comments[0].RichText.Add("x");
                var ws2 = p.Workbook.Worksheets.Add("Copied S3", ws);
                ws.InsertRow(2, 1);
                ws.DeleteRow(2);
                ws2.DeleteRow(2);
                ws.InsertRow(2, 2);
                var c = ws2.Comments;   // Access the comment collection to force loading it. Otherwise Exception!
                int dummy = c.Count;    // to load!
                p.Workbook.Worksheets.Delete(ws);
                p.Workbook.Worksheets.Delete(ws2);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void DeleteWorksheetIssue()
        {
            using (var p = OpenTemplatePackage("CommentDelete.xlsx"))
            {
                var ws = p.Workbook.Worksheets["S3"];
                var c = ws.Comments; // Access the comment collection to force loading it. Otherwise Exception!
                int dummy = c.Count; // to load!
                //ws.DeleteRow(2);
                dummy = c.Count; // to load!
                p.Workbook.Worksheets.Delete(ws);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue294()
        {
            using (var p = OpenTemplatePackage("test_excel_workbook_before2-xl.xlsx"))
            {
                var s = p.Workbook.Styles.NamedStyles.Count;
                var ws = p.Workbook.Worksheets["Summary"];
                p.Save();
            }
        }
        [TestMethod]
        public void Issue333_2()
        {
            using (var p = OpenTemplatePackage("issue333-2.xlsx"))
            {
                var sheet = p.Workbook.Worksheets[1];
                Assert.IsFalse(string.IsNullOrEmpty(sheet.Cells[8, 1].Formula));
                Assert.IsFalse(string.IsNullOrEmpty(sheet.Cells[9, 1].Formula));
                Assert.IsFalse(string.IsNullOrEmpty(sheet.Cells[9, 2].Formula));
                Assert.IsFalse(string.IsNullOrEmpty(sheet.Cells[32, 2].Formula));
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void S107()
        {
            using (var p = OpenTemplatePackage("2021-03-18 - Styling issues.xlsm"))
            {
                p.Save();
                var p2 = new ExcelPackage(p.Stream);
                p2.SaveAs(new FileInfo(p.File.DirectoryName + "\\Test.xlsm"));
            }
        }
        [TestMethod]
        public void S127()
        {
            using (var p = OpenTemplatePackage("Tagging Template V15 - New Format.xlsx"))
            {
                SaveWorkbook("Tagging Template V15 - New Format2.xlsx", p);
            }
        }
        [TestMethod]
        public void MergeIssue()
        {
            using (var p = OpenTemplatePackage("MergeIssue.xlsx"))
            {
                var ws = p.Workbook.Worksheets["s7"];
                ws.Cells["B2:F2"].Merge = false;
                ws.Cells["B2:F12"].Clear();
                ws.Cells["B2:F2"].Merge = true;

                ws.Cells["B2:F12"].Merge = false;
                ws.Cells["B2:F12"].Clear();

                ws.Cells["B2:F12"].Merge = true;
                ws.Cells["B1:F12"].Clear();

                ws.Cells["B2:F2"].Merge = true;
                ws.Cells["B2:F2"].Merge = false;
                ws.Cells["B2:F12"].Clear();
                ws.Cells["B2:F2"].Merge = true;
            }
        }

        public void DefinedNamesAddressIssue()
        {
            using (var p = OpenPackage("defnames.xlsx"))
            {
                var ws1 = p.Workbook.Worksheets.Add("Sheet1");
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");

                var name = ws1.Names.Add("Name2", ws1.Cells["B1:C5"]);
                Assert.AreEqual("Sheet1", name.Worksheet.Name);
                name.Address = "Sheet3!B2:C6";
                Assert.IsNull(name.Worksheet);
                Assert.AreEqual("Sheet3", name.WorkSheetName);

            }
        }
        [TestMethod]
        public void Issue341()
        {
            using (var package = OpenTemplatePackage("Base_debug.xlsx"))
            {
                using (var atomic_sheet_package = OpenTemplatePackage("Test_debug.xlsx"))
                {
                    var s = atomic_sheet_package.Workbook.Worksheets["Test3"];
                    var s_copy = package.Workbook.Worksheets.Add("Test3", s); // Exception on this line
                    s_copy.Drawings[0].As.Chart.LineChart.Series[0].XSeries = "A1:a15";
                    atomic_sheet_package.Save();
                }
                package.Save();
            }
        }
        [TestMethod]
        public void Issue347_2()
        {
            using (var p = OpenTemplatePackage("i347.xlsx"))
            {
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue353()
        {
            using (var p = OpenTemplatePackage("HeaderFooterTest (1).xlsx"))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets[0];
                Assert.IsFalse(worksheet.HeaderFooter.differentFirst);
                Assert.IsFalse(worksheet.HeaderFooter.differentOddEven);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue354()
        {
            using (var p = OpenTemplatePackage("i354.xlsx"))
            {
                var ws1 = p.Workbook.Worksheets[0];
                var ws2 = p.Workbook.Worksheets[2];
                var pt = ws1.PivotTables.Add(ws1.Cells["A2"], ws2.Cells["A1:E3005"], "pt");
                ws2.Cells["B2"].Value = eDateGroupBy.Years;
                ws2.Cells["B3"].Value = eDateGroupBy.Months;
                pt.ColumnFields.Add(pt.Fields[1]);
                pt.RowFields.Add(pt.Fields[4]);
                pt.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void StyleIssueLibreOffice()
        {
            foreach (bool onColumns in new[] { true, false })
            {
                var ep = new ExcelPackage();
                var ws = ep.Workbook.Worksheets.Add("Test");

                // Header area (along with freezing the header in the view)
                ws.Cells[1, 1, 1, 8].Style.Font.Bold = true;
                for (int i = 1; i < 9; ++i)
                    ws.Cells[1, i].Value = $"Test {i}";
                ws.View.FreezePanes(2, 1);

                if (onColumns)
                {
                    // Set the horizontal alignment on the columns themselves
                    ws.Column(3).Style.HorizontalAlignment = ws.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    ws.Column(5).Style.HorizontalAlignment = ws.Column(6).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    ws.Column(7).Style.HorizontalAlignment = ws.Column(8).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }
                else
                {
                    // Set the horizontal alignment on the cells of the header
                    ws.Cells[1, 3, 1, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    ws.Cells[1, 5, 1, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    ws.Cells[1, 7, 1, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }

                for (int row = 2; row < 30; ++row)
                {
                    for (int i = 1; i < 9; ++i)
                        ws.Cells[row, i].Value = row % 2 == 0 ? (8 * (row - 2) + i).ToString() : $"Test {8 * (row - 2) + i}";
                    if (!onColumns)
                    {
                        // Set the horizontal alignment on this row's cells
                        ws.Cells[row, 3, row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        ws.Cells[row, 5, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 7, row, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    }
                }

                ws.Cells.AutoFitColumns(0);

                ep.SaveAs(new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"AlignmentTest-On{(onColumns ? "Columns" : "Cells")}.xlsx")));
            }
        }
        [TestMethod]
        public void Issue382()
        {
            using (var p = OpenPackage("Issue382.xlsx", true))
            {
                p.Workbook.Styles.NamedStyles[0].Style.Font.Size = 9;
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Cell Value";
                ws.Cells.AutoFitColumns();
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue381()
        {
            using (var p = OpenTemplatePackage("Issue381.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                Assert.AreEqual(2, ws.Drawings.Count);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ReadIssue()
        {
            using (var p = OpenTemplatePackage("cf.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
            }
        }
        [TestMethod]
        public void Issue407_1()
        {
            using (var p = OpenTemplatePackage("TestStyles_MoreCellStyleXfsThanCellXfs.xlsx"))
            {
                var p2 = new ExcelPackage();
                var ws = p2.Workbook.Worksheets.Add("Copied Style", p.Workbook.Worksheets["StylesTestSheet"]);

                SaveWorkbook("Issue407_1.xlsx", p2);
            }
        }
        [TestMethod]
        public void Issue407_2()
        {
            using (var p = OpenTemplatePackage("TestStyles_MinimalWithNamedStyles.xlsx"))
            {
                var p2 = new ExcelPackage();
                var ws = p.Workbook.Worksheets["StylesTestSheet"];

                Assert.AreEqual("Normal", ws.Cells["A1"].StyleName);
                Assert.AreEqual("MyCustomCellStyle", ws.Cells["A2"].StyleName);
                Assert.AreEqual("Normal", ws.Cells["A3"].StyleName);
                Assert.AreEqual("MyCalculationStyle", ws.Cells["A4"].StyleName);

                Assert.AreEqual("Normal", ws.Cells["B1"].StyleName);
                Assert.AreEqual("MyBoldStyle1", ws.Cells["B2"].StyleName);
                Assert.AreEqual("MyBoldStyle2", ws.Cells["B3"].StyleName);
                ws = p2.Workbook.Worksheets.Add("Copied Style", p.Workbook.Worksheets["StylesTestSheet"]);
                Assert.AreEqual("Normal", ws.Cells["A1"].StyleName);
                Assert.AreEqual("MyCustomCellStyle", ws.Cells["A2"].StyleName);
                Assert.AreEqual("Normal", ws.Cells["A3"].StyleName);
                Assert.AreEqual("MyCalculationStyle", ws.Cells["A4"].StyleName);

                Assert.AreEqual("Normal", ws.Cells["B1"].StyleName);
                Assert.AreEqual("MyBoldStyle1", ws.Cells["B2"].StyleName);
                Assert.AreEqual("MyBoldStyle2", ws.Cells["B3"].StyleName);

                SaveWorkbook("Issue407_2.xlsx", p2);
            }
        }
        [TestMethod]
        public void s185()
        {
            using (var p = OpenTemplatePackage("s185.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var chart = ws.Drawings[0] as ExcelLineChart;

                Assert.AreEqual(4887, chart.PlotArea.ChartTypes[1].Series[0].NumberOfItems);
            }
        }
        [TestMethod]
        public void GetNamedRangeAddressAfterRowInsert()
        {
            using (var pck = OpenTemplatePackage("TestWbk_SingleNamedRange.xlsx"))
            {
                // Get the worksheet containing the named range
                var ws = pck.Workbook.Worksheets["Sheet1"];
                // Get the named range
                var namedRange = ws.Names["MyValues"];
                // Check that the named range exists with the expected address
                Assert.AreEqual("Sheet1!$A$1:$A$9", namedRange.FullAddress);
                Assert.AreEqual("Sheet1!$A$1:$A$9", namedRange.Address); // This line is currently failing
                                                                         // Insert a row in the middle of the range
                ws.InsertRow(5, 1);
                // Check that the named range's address has been correctly updated
                Assert.AreEqual("Sheet1!$A$1:$A$10", namedRange.FullAddress);
                Assert.AreEqual("Sheet1!$A$1:$A$10", namedRange.Address);
            }
        }
        [TestMethod]
        public void Issue410()
        {
            using (var package = OpenTemplatePackage("test-in.xlsx"))
            {
                var wb = package.Workbook;
                var worksheet = wb.Worksheets.Add("Pivot Tables");
                var table = wb.Worksheets[0].Tables["Table1"];
                ExcelPivotTable pt = worksheet.PivotTables.Add(worksheet.Cells["A1"], table, "PT1");
                pt.RowFields.Add(pt.Fields["ColC"]);
                pt.DataFields.Add(pt.Fields["ColB"]);
                SaveWorkbook("test-out.xlsx", package);
            }
        }
        [TestMethod]
        public void Issue415()
        {
            using (var package = OpenTemplatePackage("Issue415.xlsm"))
            {
                var wb = package.Workbook;
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void Issue417()
        {
            using (var package = OpenTemplatePackage("Issue417.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
                Assert.AreEqual("0", ws.Cells["A1"].Text);
                Assert.AreEqual(null, ws.Cells["A2"].Text);
            }
        }
        [TestMethod]
        public void Issue395()
        {
            using (var package = OpenTemplatePackage("Issue395.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void Issue418()
        {
            using (var p = OpenPackage("issue418.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Test");

                var mergetest = ws.Cells[2, 1, 2, 5];
                mergetest.IsRichText = true;
                mergetest.Merge = true;
                mergetest.Style.WrapText = true;

                var t1 = mergetest.RichText.Add($"Text 1", true);
                t1.Size = 16;
                t1.Bold = true;

                var t2 = mergetest.RichText.Add($"Text 2", true);
                t2.Size = 12;
                t2.Bold = false;

                ws.Row(2).Height = 50;

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void IssueHidden()
        {
            using (var p = OpenTemplatePackage("workbook.xlsx"))
            {
                p.Workbook.Worksheets[p.Workbook.Worksheets.Count - 2].Hidden = eWorkSheetHidden.Hidden;
                p.Workbook.Worksheets[p.Workbook.Worksheets.Count - 1].Hidden = eWorkSheetHidden.Hidden;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue430()
        {
            using (var p = OpenTemplatePackage("issue430.xlsx"))
            {
                var workbook = p.Workbook;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue435()
        {
            using (var p = OpenTemplatePackage("issue435.xlsx"))
            {
                var workbook = p.Workbook;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void VbaIssueLoad()
        {
            using (var p = OpenTemplatePackage("PlantillaDefectivo-NotWorking.xlsm"))
            {
                var workbook = p.Workbook;
                var vba = p.Workbook.VbaProject;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue440()
        {
            using (var p = OpenTemplatePackage("issue440.xlsx"))
            {
                var wb = p.Workbook;
                var worksheet = wb.Worksheets.Add("Pivot Tables");
                var table = wb.Worksheets[0].Tables["Table1"];
                ExcelPivotTable pt = worksheet.PivotTables.Add(worksheet.Cells["A1"], table, "PT1");
                pt.RowFields.Add(pt.Fields["ColC"]);
                pt.DataFields.Add(pt.Fields["ColB"]);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue441()
        {
            using (var pck = OpenPackage("issue441.xlsx", true))
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var commentAddress = "B2";
                wks.Comments.Add(wks.Cells[commentAddress], "This is a comment.", "author");
                wks.Cells[commentAddress].Value = "This cell contains a comment.";

                wks.Cells["B1:B3"].Insert(eShiftTypeInsert.Right);
                commentAddress = "C2";
                Assert.AreEqual(1, wks.Comments.Count);
                Assert.AreEqual("This is a comment.", wks.Comments[0].Text);
                Assert.AreEqual("This cell contains a comment.", wks.Cells[commentAddress].GetValue<string>());
                Assert.AreEqual(commentAddress, wks.Comments[0].Address);
                SaveAndCleanup(pck);
            }
        }
        [TestMethod]
        public void Issue442()
        {
            using (var pck = OpenPackage("issue442.xlsx", true))
            {
                // Add a sheet with data validation in cell B2
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var dataValidationList = wks.DataValidations.AddListValidation("B2");
                var values = dataValidationList.Formula.Values;
                values.Add("Yes");
                values.Add("No");
                wks.Cells["B2"].Value = "Yes";

                // Confirm this was added in the right place
                Assert.AreEqual("Yes", wks.Cells["B2"].GetValue<string>());
                Assert.AreEqual("B2", wks.DataValidations[0].Address.Address);

                // Insert cells to shift this data validation to the right
                wks.Cells["B2:C5"].Insert(eShiftTypeInsert.Right);

                // Check the data validation has been moved to the right place
                Assert.AreEqual("Yes", wks.Cells["D2"].GetValue<string>());
                Assert.AreEqual("D2", wks.DataValidations[0].Address.Address);

                SaveAndCleanup(pck);
            }
        }
        [TestMethod]
        public void s224()
        {
            using (var p = OpenTemplatePackage("s224.xltx"))
            {
                SaveWorkbook("s224.xlsx", p);
            }
        }
        [TestMethod]
        public void InsertCellsNextToComment()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");

                // Create a comment in cell B2
                var commentAddress = "B2";
                wks.Comments.Add(wks.Cells[commentAddress], "This is a comment.", "author");
                wks.Cells[commentAddress].Value = "This cell contains a comment.";

                // Create another comment in cell B10
                var commentAddress2 = "B10";
                wks.Comments.Add(wks.Cells[commentAddress2], "This is another comment.", "author");
                wks.Cells[commentAddress2].Value = "This cell contains another comment.";

                // Insert cells so the first comment is now in C2
                wks.Cells["B1:B3"].Insert(eShiftTypeInsert.Right);
                commentAddress = "C2";

                // Check that both the cell value and the comment was correctly moved
                Assert.AreEqual(2, wks.Comments.Count);
                Assert.AreEqual("This is a comment.", wks.Comments[0].Text);
                Assert.AreEqual("This cell contains a comment.", wks.Cells[commentAddress].GetValue<string>());
                Assert.AreEqual(commentAddress, wks.Comments[0].Address);

                // Check that the second comment hasn't moved
                Assert.AreEqual("This is another comment.", wks.Comments[1].Text);
                Assert.AreEqual("This cell contains another comment.", wks.Cells[commentAddress2].GetValue<string>());
                Assert.AreEqual(commentAddress2, wks.Comments[1].Address);

            }
        }
        [TestMethod]
        public void InsertRowsIntoTable_CheckFormulasWithColumnReferences()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet, and a table with headers and a single row of data
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                wks.Cells["B2:D3"].Value = new object[,] { { "Col1", "Col2", "Col3" }, { 1, 2, 3 } };
                wks.Tables.Add(wks.Cells["B2:D3"], "Table1");

                // Add a SUM formula on the worksheet with a reference to a table column
                wks.Cells["C10"].Formula = "SUM(Table1[Col2])";
                wks.Cells["C10"].Calculate();

                Assert.AreEqual(2.0, wks.Cells["C3"].GetValue<double>());
                Assert.AreEqual("SUM(Table1[Col2])", wks.Cells["C10"].Formula);
                Assert.AreEqual(2.0, wks.Cells["C10"].GetValue<double>());

                // Insert 2 rows into the worksheet to extend the table
                wks.InsertRow(4, 2);

                // Check that the formula that was in C10 (now C12) still references the column
                Assert.AreEqual("SUM(Table1[Col2])", wks.Cells["C12"].Formula);
            }
        }
        [TestMethod]
        public void CopyWorksheetWithDynamicArrayFormula()
        {
            using (var p1 = OpenTemplatePackage("TestDynamicArrayFormula.xlsx"))
            {
                using (var p2 = new ExcelPackage())
                {
                    p2.Workbook.Worksheets.Add("Sheet1", p1.Workbook.Worksheets["Sheet1"]);
                    SaveWorkbook("DontCopyMetadataToNewWorkbook.xlsx", p2);
                }
            }

            Assert.Inconclusive("Now try to open the file in Excel - does it tell you the file is corrupt?");
        }
        [TestMethod]
        public void VbaIssue()
        {
            using (var p = OpenTemplatePackage("Issue479.xlsm"))
            {
                p.Workbook.Worksheets.Add("New Sheet");
                SaveAndCleanup(p);
            }
        }
        public static readonly Dictionary<string, string> ASSET_FIELDS = new Dictionary<string, string> { { string.Empty, "Select one..." }, { "APPRAISAL_DATE", "Appraisal Date" }, { "APPRAISAL_AREA", "Appraisal Surface" }, { "APPRAISAL_AREA_CCAA", "Appraisal Surface w/CCAA" }, { "APPRAISAL_VALUE", "Appraisal Value" }, { "AURA" + "." + "REFERENCE", "Aura ID" }, { "BATHROOMS", "Bathroom" }, { "BORROWER_ID", "Borrower ID" }, { "AREA_CCAA", "Built surface w/CCAA" }, { "CADASTRAL_REFERENCE", "Cadastral reference" }, { "DEVELOPMENT" + "." + "CLIENT_GROUP_ID", "Client Development ID" }, { "CLIENT_ID", "Client ID" }, { "YEAR_OF_CONSTRUCTION", "Construction Year" }, { "COUNTRY", "Country" }, { "CROSSING_DOCKS", "Crossing Docks" }, { "DATE_OF_DISQUALIFICATION", "Date of desaffection - Social Housing" }, { "LEADER" + "." + "REFERENCE", "Dependency Reference" }, { "DEVELOPMENT" + "." + "REFERENCE", "Development ID" }, { "ADDRESS_DOOR", "Door" }, { "DUPLEX", "Duplex" }, { "ELEVATOR", "Elevator" }, { "ADDRESS_FLOOR", "Floor" }, { "FULL_ADDRESS", "Full Address" }, { "IDUFIR", "IDUFIR" }, { "ILLEGAL_SQUATTERS", "Illegal Squatters" }, { "ORIENTATION", "Interior/Exterior" }, { "LATITUDE", "Latitude" }, { "LIEN", "Lien" }, { "LOAN_ID", "Loan ID" }, { "LONGITUDE", "Longitude" }, { "MAINTENANCE_STATUS", "Maintenance Status" }, { "MARKET_SHARE", "Market Share (%)" }, { "LEGAL_MAXIMUM_VALUE", "Max. Value - Social Housing" }, { "MAX_HEIGHT", "Maximum Height" }, { "VPO_MODULE", "Module - Social Housing" }, { "MUNICIPALITY", "Municipality" }, { "NEGATIVE_COLD", "Negative Cold" }, { "ADDRESS_NUMBER", "Number" }, { "BORROWER", "Owner" }, { "PARKINGS", "Parking" }, { "PERIMETER", "Perimeter" }, { "PLOT_AREA", "Plot Surface" }, { "POSITIVE_COLD", "Positive Cold" }, { "PROVINCE", "Province" }, { "REFERENCE", "Reference ID" }, { "REGISTRATION", "Registry" }, { "REGISTRY_ID", "Registry ID" }, { "REGISTRATION_NUMBER", "Registry Number" }, { "AREA_REGISTRY", "Registry Surface" }, { "AREA_CCAA_REGISTRY", "Registry Surface w/CCAA" }, { "RENTED", "Rented" }, { "REPEATED", "Repeated" }, { "ROOMS", "Rooms" }, { "SCOPE", "Scope" }, { "SEA_VIEWS", "Sea Views" }, { "MONTHLY_COMM_EXP_SQM", "Service Charges" }, { "SMOKE_VENT", "Smoke Ventilation" }, { "VPO", "Social Housing" }, { "DEVELOPMENT" + "." + "PROPERTY_STATUS", "Status" }, { "STOREROOMS", "Storage" }, { "ADDRESS_NAME", "Street" }, { "ASSET_SUBTYPE", "Sub-typology" }, { "AREA", "Surface" }, { "SWIMMING_POOL", "Swimming Pool" }, { "TERRACE", "Terrace" }, { "TERRACE_AREA", "Terrace Surface" }, { "ACTIVITY", "Type of activity" }, { "STATE", "Type of product" }, { "ASSET_TYPE", "Typology" }, { "USEFUL_AREA", "Useful Surface" }, { "VALUATION_TYPE", "Valuation Type" }, { "ZIP_CODE", "Zip Code" } };

        public class Error { public string TypeOfError { get; set; } public int Row { get; set; } public int Col { get; set; } public List<string> Messages { get; set; } }

        public class AssetField { public int Index { get; set; } public string Field { get; set; } }

        [TestMethod]
        public void Issue478()
        {

            var dataStartRow = 2;
            var errors = JsonConvert.DeserializeObject<Error[]>("[{\"typeOfError\":\"WARNING\",\"row\":4,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":20,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":35,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":47,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":57,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":60,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":90,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":131,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":136,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":138,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":139,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]}]");
            var assetFields = JsonConvert.DeserializeObject<AssetField[]>("[{\"index\":1,\"field\":\"Reference\"},{\"index\":15,\"field\":\"ZipCode\"},{\"index\":16,\"field\":\"Municipality\"},{\"index\":17,\"field\":\"FullAddress\"}]");

            using (var excelPackage = OpenTemplatePackage("issue478.xlsx"))
            {
                var worksheet = excelPackage.Workbook.Worksheets["Avances"];
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                // Add column of errors and warnings
                var startMessagesColumn = end.Column + 1;
                worksheet.InsertColumn(startMessagesColumn, 2);
                var errorColumn = startMessagesColumn;
                var warningColumn = startMessagesColumn + 1;
                worksheet.Cells[(dataStartRow) - 1, errorColumn].Value = "Errors";
                worksheet.Cells[(dataStartRow) - 1, warningColumn].Value = "Warnings";
                foreach (var error in errors)
                {
                    if (error.TypeOfError == "ERROR")
                    {
                        //worksheet.Cells[error.Row - 1, errorColumn].Value += string.Join(" ", error.Messages.Select(w => string.Format("{0} {1}", ASSET_FIELDS.GetValueOrDefault(assetFields.Where(x => x.Index == error.Col).Select(x => x.Field).FirstOrDefault()), w)));
                    }
                    else
                    {
                        //worksheet.Cells[error.Row - 1, warningColumn].Value += string.Join(" ", error.Messages.Select(w => string.Format("{0} {1}", ASSET_FIELDS.GetValueOrDefault(assetFields.Where(x => x.Index == error.Col).Select(x => x.Field).FirstOrDefault()), w)));
                    }
                }

                // Remove distinct columns from "Reference"
                var colFieldReference = assetFields.Where(x => x.Field == "REFERENCE").Select(x => x.Index).FirstOrDefault();
                worksheet.Cells[1, colFieldReference + 1].Value = "Reference";

                var deletedColumns = 0;
                for (int i = 1; i <= end.Column; i++)
                {
                    if (colFieldReference + 1 != i && errorColumn != i && warningColumn != i)
                    {
                        worksheet.DeleteColumn(i - deletedColumns);
                        deletedColumns++;
                    }
                }

                // Remove rows that do not contain errors
                var deletedRows = 0;
                for (int i = 1; i <= end.Row; i++)
                {
                    if (i < (dataStartRow - 1) || (i >= dataStartRow && !errors.Any(w => (w.Row - 1) == i)))
                    {
                        worksheet.DeleteRow(i - deletedRows);
                        deletedRows++;
                    }
                }
                SaveAndCleanup(excelPackage);
            };
        }
        [TestMethod]
        public void TestColumnWidthsAfterDeletingColumn()
        {
            using (var pck = OpenTemplatePackage("Issue480.xlsx"))
            {
                // Get the worksheet where columns 3-5 have a width of around 18
                var wks = pck.Workbook.Worksheets["Sheet1"];

                // Check the width of column 5
                Assert.AreEqual(18.77734375, wks.Column(5).Width, 1E-5);

                // Delete column 4
                wks.DeleteColumn(4, 3);

                // Check width of column 5 (now 4) hasn't changed
                Assert.AreEqual(18.77734375, wks.Column(3).Width, 1E-5);
                //Assert.AreEqual(18.77734375, wks.Column(4).Width, 1E-5);
            }
        }

        [TestMethod]
        public void Issue484_InsertRowCalculatedColumnFormula()
        {
            using (var p = new ExcelPackage())
            {
                // Create some worksheets
                var ws1 = p.Workbook.Worksheets.Add("Sheet1");

                // Create some tables with calculated column formulas
                var tbl1 = ws1.Tables.Add(ws1.Cells["A11:C12"], "Table1");
                tbl1.Columns[2].CalculatedColumnFormula = "A12+B12";

                var tbl2 = ws1.Tables.Add(ws1.Cells["E11:G12"], "Table2");
                tbl2.Columns[2].CalculatedColumnFormula = "A12+F12";

                // Check the formulas have been set correctly
                Assert.AreEqual("A12+B12", ws1.Cells["C12"].Formula);
                Assert.AreEqual("A12+F12", ws1.Cells["G12"].Formula);
                Assert.AreEqual("A12+B12", tbl1.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("A12+F12", tbl2.Columns["Column3"].CalculatedColumnFormula);

                // Delete two rows above the tables
                ws1.DeleteRow(5, 2);

                // Check the formulas were updated
                Assert.AreEqual("A10+B10", ws1.Cells["C10"].Formula);
                Assert.AreEqual("A10+F10", ws1.Cells["G10"].Formula);
                Assert.AreEqual("A10+B10", tbl1.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("A10+F10", tbl2.Columns[2].CalculatedColumnFormula);
            }
        }
        [TestMethod]
        public void Issue484_DeleteRowCalculatedColumnFormula()
        {
            using (var p = new ExcelPackage())
            {
                // Create some worksheets
                var ws1 = p.Workbook.Worksheets.Add("Sheet1");

                // Create some tables with calculated column formulas
                var tbl1 = ws1.Tables.Add(ws1.Cells["A11:C12"], "Table1");
                tbl1.Columns[2].CalculatedColumnFormula = "A12+B12";

                var tbl2 = ws1.Tables.Add(ws1.Cells["E11:G12"], "Table2");
                tbl2.Columns[2].CalculatedColumnFormula = "A12+F12";

                // Check the formulas have been set correctly
                Assert.AreEqual("A12+B12", ws1.Cells["C12"].Formula);
                Assert.AreEqual("A12+F12", ws1.Cells["G12"].Formula);
                Assert.AreEqual("A12+B12", tbl1.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("A12+F12", tbl2.Columns["Column3"].CalculatedColumnFormula);

                // Delete two rows above the tables
                ws1.DeleteRow(5, 2);

                // Check the formulas were updated
                Assert.AreEqual("A10+B10", ws1.Cells["C10"].Formula);
                Assert.AreEqual("A10+F10", ws1.Cells["G10"].Formula);
                Assert.AreEqual("A10+B10", tbl1.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("A10+F10", tbl2.Columns[2].CalculatedColumnFormula);
            }
        }
        [TestMethod]
        public void FreezeTemplate()
        {
            using (var p = OpenTemplatePackage("freeze.xlsx"))
            {
                // Get the worksheet where columns 3-5 have a width of around 18
                var ws = p.Workbook.Worksheets[0];
                ws.View.FreezePanes(40, 5);
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void CopyWorksheetWithBlipFillObjects()
        {
            using (var p1 = OpenTemplatePackage("BlipFills.xlsx"))
            {
                var ws = p1.Workbook.Worksheets[0];
                var wsCopy = p1.Workbook.Worksheets.Add("Copy", p1.Workbook.Worksheets[0]);

                ws.Cells["G4"].Copy(wsCopy.Cells["F20"]);
                SaveAndCleanup(p1);
            }
        }
        [TestMethod]
        public void Issue519()
        {
            using (var package = OpenPackage("I519.xlsx", true))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[3, 1].Value = "test";

                var ctrl = worksheet.Drawings.AddCheckBoxControl("test");
                ctrl.SetPosition(10, 10);
                ctrl.Checked = OfficeOpenXml.Drawing.Controls.eCheckState.Checked; // creates valid XLSX file
                //ctrl.Checked = OfficeOpenXml.Drawing.Controls.eCheckState.Mixed; // creates valid XLSX file
                //ctrl.Checked = OfficeOpenXml.Drawing.Controls.eCheckState.Unchecked; // creates invalid XLSX file

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void Issue520_2()
        {
            using (var p = OpenPackage("i520.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("17Columns");
                var tableData = Enumerable.Range(1, 10)
                .Select(_ => new
                {
                    C01 = 1,
                    C02 = 2,
                    C03 = 3,
                    C04 = 4,
                    C05 = 5,
                    C06 = 6,
                    C07 = 7,
                    C08 = 8,
                    C09 = 9,
                    C10 = 10,
                    C11 = 11,
                    C12 = 12,
                    C13 = 13,
                    C14 = 14,
                    C15 = 15,
                    C16 = 16,
                    C17 = 17
                }).ToArray();
                var table = sheet.Cells[1, 1].LoadFromCollection(tableData, true, TableStyles.Light1);
                table.AutoFitColumns();

                sheet = p.Workbook.Worksheets.Add("16Columns");
                var tableData2 = Enumerable.Range(1, 10)
                .Select(_ => new
                {
                    C01 = 1,
                    C02 = 2,
                    C03 = 3,
                    C04 = 4,
                    C05 = 5,
                    C06 = 6,
                    C07 = 7,
                    C08 = 8,
                    C09 = 9,
                    C10 = 10,
                    C11 = 11,
                    C12 = 12,
                    C13 = 13,
                    C14 = 14,
                    C15 = 15,
                    C16 = 16
                }).ToArray();
                table = sheet.Cells[1, 1].LoadFromCollection(tableData2, true, TableStyles.Light1);
                table.AutoFitColumns();

                sheet = p.Workbook.Worksheets.Add("18Columns");
                var tableData3 = Enumerable.Range(1, 10)
                .Select(_ => new
                {
                    C01 = 1,
                    C02 = 2,
                    C03 = 3,
                    C04 = 4,
                    C05 = 5,
                    C06 = 6,
                    C07 = 7,
                    C08 = 8,
                    C09 = 9,
                    C10 = 10,
                    C11 = 11,
                    C12 = 12,
                    C13 = 13,
                    C14 = 14,
                    C15 = 15,
                    C16 = 16,
                    C17 = 17,
                    C18 = 18
                }).ToArray();
                table = sheet.Cells[1, 1].LoadFromCollection(tableData3, true, TableStyles.Light1);
                table.AutoFitColumns();
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void Issue522()
        {
            using (var package = OpenPackage("I22.xlsx", true))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[1, 1].Value = -1234;
                worksheet.Cells[1, 1].Style.Numberformat.Format = "#.##0\"*\";(#.##0)\"*\"";
                var s = worksheet.Cells[1, 1].Text;

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void IssueNamedRanges()
        {
            using (var package = OpenTemplatePackage("ORRange23 Problem.xlsx"))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[1, 1].Value = -1234;
                worksheet.Cells[1, 1].Style.Numberformat.Format = "#.##0\"*\";(#.##0)\"*\"";
                var s = worksheet.Cells[1, 1].Text;

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void DvcfCopy()
        {
            using (var p = OpenTemplatePackage("i527.xlsm"))
            {

                // Fails when data validation is set
                // Fails when conditional formatting is set.
                var copyFrom1 = p.Workbook.Worksheets["CopyFrom"].Cells["A1:BR23"];
                var copyTo1 = p.Workbook.Worksheets["CopyTo"].Cells["A:XFD"];
                copyFrom1.Copy(copyTo1);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s268()
        {
            using (var p = OpenTemplatePackage("s268.xlsx"))
            {
                var s3 = p.Workbook.Worksheets["s3"];

                s3.InsertRow(1, 1);
                s3.InsertRow(1, 1);
                s3.InsertRow(1, 1);
                s3.InsertRow(1, 1);
                s3.InsertRow(1, 1);
                s3.InsertRow(1, 1);
                s3.InsertRow(1, 1);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void Issue538()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                var validation = sheet.DataValidations.AddListValidation("A1 B1");
                validation.Formula.ExcelFormula = "Sheet2!$A$7:$A$12"; // throws exception "Multiple addresses may not be commaseparated, use space instead"
            }
        }
        [TestMethod]
        public void s272()
        {
            using (var p = OpenTemplatePackage("RadioButton.xlsm"))
            {
                if (p.Workbook.VbaProject == null)
                {
                    p.Workbook.CreateVBAProject();
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s277()
        {
            using (var p = OpenTemplatePackage("s277.xlsx"))
            {
                foreach (var ws in p.Workbook.Worksheets)
                    ws.Drawings.Clear();
            }
        }
        [TestMethod]
        public void s279()
        {
            using (var p = OpenTemplatePackage("s279.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.Cells["C3"].Value = "Test";
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void I546()
        {
            using (var excelPackage = OpenTemplatePackage("b.xlsx"))
            {
                var ws = excelPackage.Workbook.Worksheets[0];
                var cell = ws.Cells["A2"];
                var formula = cell.Formula;
                var value1 = cell.Value;
                Console.WriteLine($"value1: {value1}");

                var externalLinks = excelPackage.Workbook.ExternalLinks;
                var externalWorkbook = externalLinks[0].As.ExternalWorkbook;
                externalWorkbook.Load();

                ws.ClearFormulaValues();
                ws.Calculate(); // "Circular reference occurred at A2" exception is thrown here

                var value2 = cell.Value;
                Console.WriteLine($"value2: {value2}");
            }
        }
        [TestMethod]
        public void I548()
        {
            using (var p = OpenTemplatePackage("09-145.xlsx"))
            {
                var wsCopy = p.Workbook.Worksheets["Sheet3"];
                var ws = p.Workbook.Worksheets.Add("tmpCopy");
                //copy in the same o in another workbook, same issue
                wsCopy.Cells["C1:AB55"].Copy(ws.Cells["C1"], ExcelRangeCopyOptionFlags.ExcludeFormulas);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void I552()
        {
            using (var package = OpenTemplatePackage("I552-2.xlsx"))
            {
                var worksheet = package.Workbook.Worksheets[0];
                worksheet.InsertRow(2, 1);
                worksheet.Cells[1, 1, 1, 10].Copy(worksheet.Cells[2, 1, 2, 10]);

                SaveAndCleanup(package);
            }

            using (var package = OpenPackage("I552-2.xlsx"))
            {
                var worksheet = package.Workbook.Worksheets[0];
                worksheet.InsertRow(2, 1);
                worksheet.Cells[1, 1, 1, 10].Copy(worksheet.Cells[2, 1, 2, 10]);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s285()
        {
            using (var package = OpenTemplatePackage("s285.xlsx"))
            {
                var worksheet = package.Workbook.Worksheets[0];
                worksheet.SetValue(3, 3, "Test");
                var ns = package.Workbook.Styles.CreateNamedStyle("Normal");
                ns.BuildInId = 0;
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void i566()
        {
            using (var package = OpenPackage("i566.xlsx", true))
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet 1");
                var ws = package.Workbook.Worksheets["Sheet 1"];
                ws.SetValue(3, 3, "Test");
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void i583()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.SetValue(1048576, 1, 1);
                Assert.AreEqual("A1048576", ws.Dimension.Address);

            }
        }

        [TestMethod]
        public void i567()
        {
            using (var package = OpenTemplatePackage("i567.xlsx"))
            {
                var wsSource = package.Workbook.Worksheets["Detail"];

                var dataCollection = new List<object[]>()
                {
                    new object[]{"Driver 1",1,2,"Fleet 1", "Manager 1",3,true,0,5,0 },
                    new object[]{"Driver 2",3,4,"Fleet 2", "Manager 2", 5,true,0,8,0 }
                };
                wsSource.Cells["A1"].Value = null;
                //code to load a collection to the spreadsheet. very nice
                wsSource.Cells["A2"].LoadFromArrays(dataCollection);

                foreach (var ws in package.Workbook.Worksheets)
                {
                    foreach (var pt in ws.PivotTables)
                    {
                        pt.CacheDefinition.SourceRange = wsSource.Cells["A1:J3"];
                    }
                }

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void i574()
        {
            using (var package = OpenTemplatePackage("i574.xlsx"))
            {
                var wsSource = package.Workbook.Worksheets[0];

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void LoadFontSize()
        {
            FontSize.LoadAllFontsFromResource();
        }

        [TestMethod]
        public void PiechartWithHorizontalSource()
        {
            using (var p = OpenPackage("piechartHorizontal.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("PieVertical");
                ws.SetValue("A1", "C1");
                ws.SetValue("A2", "C2");
                ws.SetValue("A3", "C3");
                ws.SetValue("B1", 15);
                ws.SetValue("B2", 45);
                ws.SetValue("B3", 40);

                var chart = ws.Drawings.AddPieChart("Pie1", ePieChartType.Pie);
                chart.VaryColors = true;
                chart.Series.Add("B1:B3", "A1:A3");
                chart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.PieChartStyle1);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void PiechartWithVerticalSource()
        {
            using (var p = OpenPackage("piechartvertical.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("PieVertical");
                ws.SetValue("A1", "C1");
                ws.SetValue("B1", "C2");
                ws.SetValue("C1", "C3");
                ws.SetValue("A2", 15);
                ws.SetValue("B2", 45);
                ws.SetValue("C2", 40);

                var chart = ws.Drawings.AddPieChart("Pie1", ePieChartType.Pie);
                chart.VaryColors = true;
                chart.Series.Add("A2:C2", "A1:C1");
                chart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.PieChartStyle1);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void CheckEnvironment()
        {
            System.Drawing.Graphics.FromHwnd(IntPtr.Zero);
        }
        [TestMethod]
        public void Issue592()
        {
            using (var p = OpenTemplatePackage("I592.xlsx"))
            {        
                var ws = p.Workbook.Worksheets[0];
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void I594()
        {
            using (var p = OpenTemplatePackage("i594.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                var tbl = ws.Tables[0];
                tbl.AddRow(2);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s302()
        {
            using (var p = OpenTemplatePackage("SampleData.xlsx"))
            {
                var worksheet = p.Workbook.Worksheets[2];
                ExcelBarChart chart = worksheet.Drawings.AddBarChart("NewBarChart", eBarChartType.BarClustered);

                chart.SetPosition(32, 0, 1, 0);

                chart.SetSize(785, 320);

                chart.RoundedCorners = false;

                chart.Border.Fill.Color = Color.Gray;

                chart.Legend.Position = eLegendPosition.Bottom;



                ExcelBarChartSerie eventS1Serie = chart.Series.Add("D9:D12", "B9:B12");

                eventS1Serie.Header = "STATISTIQUES COMPARATIVES";

                ExcelBarChartSerie eventS2Serie = chart.Series.Add("H9:H12", "B9:B12");

                

                chart.StyleManager.SetChartStyle(ePresetChartStyle.BarChartStyle5, ePresetChartColors.MonochromaticPalette5);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void I596()
        {
            using (var p = OpenTemplatePackage("I596.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void I606()
        {
            using (var p = OpenTemplatePackage("i606.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s312()
        {
            using (var p = OpenTemplatePackage("richtext.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var t = ws.Cells["C2"].RichText.GetType(); ;
                var prop = t.GetProperty("TopNode", BindingFlags.GetProperty | BindingFlags.NonPublic | BindingFlags.Instance);
                var topNode = prop.GetValue(ws.Cells["C2"].RichText);
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void Issue308()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                package.Compatibility.IsWorksheets1Based = false;

                var wb = package.Workbook;
                wb.Worksheets.Add("A");
                wb.Worksheets.Add("B");

                ExcelWorksheet sheetA = wb.Worksheets["A"];
                ExcelWorksheet sheetB = wb.Worksheets["B"];

                var qsUndefined = QStr("Undefined");
                var qsEmpty = QStr("");
                var qsOk = QStr("OK");
                var qsError = QStr("ERROR");
                var qsES4 = QStr("ES4");
                var qsES5 = QStr("ES5");

                var errorFormula = $"IF(OR(LEFT(RC4,9)={qsUndefined},AND(OR(RC5={qsES4},RC5={qsES5}),RC8={qsEmpty}),RC5={qsEmpty}),{qsError},{qsOk})";

                sheetB.Cells[1, 4].Value = "ABC";
                sheetB.Cells[1, 5].Value = "ES1";
                sheetB.Cells[1, 8].Value = "ES1";
                sheetB.Cells[1, 10].FormulaR1C1 = errorFormula;

                wb.Calculate();

                Console.WriteLine($"Sht B Value = {sheetB.Cells[1, 10].Value} FormulaR1c1={sheetB.Cells[1, 10].FormulaR1C1}");
                Console.WriteLine($"Sht A Value = {sheetA.Cells[1, 10].Value} FormulaR1c1={sheetA.Cells[1, 10].FormulaR1C1}");

                Assert.AreEqual("OK", sheetB.Cells["J1"].Value);

                // package.SaveAs(@"c:\temp\eppTest306.xlsx");
            }
        }

        string QStr(string s)
        {
            char quotechar = '\"';
            return $"{quotechar}{s}{quotechar}";
        }
        [TestMethod]
        public void s314()
        {
            using (var p = OpenTemplatePackage("SlicerIssue.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var pt = ws.PivotTables[0];
                var wsTable = p.Workbook.Worksheets[1];
                var tbl = wsTable.Tables[0];
                wsTable.Cells["E9"].Value = "New Value";
                wsTable.Cells["E9"].Value="Stockholm";
                wsTable.Cells["F11"].Value = "Test";
                tbl.AddRow(11);
                tbl.Range.Offset(1, 0, tbl.Range.Rows - 1, tbl.Range.Columns).Copy(tbl.Range.Offset(11,0));
                Assert.IsNotNull(pt.Fields[2].Items.Count);
                Assert.IsNotNull(pt.Fields[10].Items.Count);
                Assert.IsNotNull(pt.Fields[1].Items.Count);
                for (int r=12;r<23;r++)
                {
                    wsTable.Cells[r, 1].Value += "-" + r;
                    wsTable.Cells[r, 2].Value = r;
                    wsTable.Cells[r, 3].Value += "-" + r;
                }

                wsTable.Cells[15, 1].Value = null;
                wsTable.Cells[14, 2].Value = null;
                wsTable.Cells[13, 3].Value = null;

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void i620()
        {
            using (var p = OpenTemplatePackage("i621.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.DeleteColumn(1, 3);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void i609()
        {
            using (var p = OpenTemplatePackage("i609.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void HeaderFooterWithInitWhiteSpace()
        {
            using (var p = OpenPackage("i631.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.HeaderFooter.EvenFooter.RightAlignedText = "  Row1\r\nRow 2 ";
                ws.HeaderFooter.OddFooter.RightAlignedText = "\r\nRow1\r\nRow 2\r\n";
                ws.HeaderFooter.EvenHeader.LeftAlignedText = "\tRow1\r\nRow 2";
                ws.HeaderFooter.OddHeader.LeftAlignedText = " Row1\r\nRow 2";
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void CopyDxfs()
        {
            using (var p = OpenTemplatePackage("Input Sheet.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                using (var p2 = OpenPackage("CopyDxfs.xlsx", true))
                {
                    p2.Workbook.Worksheets.Add("Sheet1", ws);
                    SaveAndCleanup(p2);
                }
            }
        }
    }
}
