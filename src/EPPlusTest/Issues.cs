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
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.Style;
using System.Data;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Drawing.Chart;
using System.Text;
using System.Dynamic;
using System.Globalization;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest
{
    /// <summary>
    /// This class contains testcases for issues on Codeplex and Github.
    /// All tests requiering an template should be set to ignored as it's not practical to include all xlsx templates in the project.
    /// </summary>
    [TestClass]
    public class Issues : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ValueFilter.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            _pck.Save();
            _pck.Dispose();
        }

        [TestInitialize]
        public void Initialize()
        {
            if (!Directory.Exists(@"c:\Temp"))
            {
                Directory.CreateDirectory(@"c:\Temp");
            }
            if (!Directory.Exists(@"c:\Temp\bug"))
            {
                Directory.CreateDirectory(@"c:\Temp\bug");
            }
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
            var path = @"C:\temp\output.xlsx";
            var file = new FileInfo(path);
            file.Delete();
            using (var ep = new ExcelPackage(file))
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
            p.SaveAs(new FileInfo(@"c:\temp\merge.xlsx"));
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
            p.SaveAs(new FileInfo(@"c:\temp\bug\copy.xlsx"));
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
            DirectoryInfo outputDir = new DirectoryInfo(@"c:\ExcelPivotTest");
            FileInfo MyFile = new FileInfo(@"c:\temp\bug\pivottable.xlsx");
            LoadData(MyFile);
            BuildPivotTable1(MyFile);
            BuildPivotTable2(MyFile);
        }

        private void LoadData(FileInfo MyFile)
        {
            if (MyFile.Exists)
            {
                MyFile.Delete();  // ensures we create a new workbook
            }

            using (ExcelPackage EP = new ExcelPackage(MyFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet wsData = EP.Workbook.Worksheets.Add("Data");
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
                EP.Save();
            }
        }
        private void BuildPivotTable1(FileInfo MyFile)
        {
            using (ExcelPackage ep = new ExcelPackage(MyFile))
            {

                var wsData = ep.Workbook.Worksheets["Data"];
                var totalRows = wsData.Dimension.Address;
                ExcelRange data = wsData.Cells[totalRows];

                var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot1");

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
                ep.Save();
                ep.Dispose();

            }

        }

        private void BuildPivotTable2(FileInfo MyFile)
        {
            using (ExcelPackage ep = new ExcelPackage(MyFile))
            {

                var wsData = ep.Workbook.Worksheets["Data"];
                var totalRows = wsData.Dimension.Address;
                ExcelRange data = wsData.Cells[totalRows];

                var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot2");

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
                ep.Save();
                ep.Dispose();

            }

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
                p.SaveAs(new FileInfo(@"c:\temp\rt.xlsx"));
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

                p.SaveAs(new FileInfo(@"c:\temp\styleerror.xlsx"));
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
                p.SaveAs(new FileInfo(@"c:\temp\rtpreserve.xlsx"));
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
                p.SaveAs(new FileInfo(@"c:\temp\activeCell.xlsx"));
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
            // Prepare
            var newFile = new FileInfo(Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx"));
            try
            {
                using (var package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add("ArrayTest");
                    ws.Cells["A1"].Value = 1;
                    ws.Cells["A2"].Value = 2;
                    ws.Cells["A3"].Value = 3;
                    ws.Cells["B1:B3"].CreateArrayFormula("A1:A3");
                    package.Save();
                }
                Assert.IsTrue(File.Exists(newFile.FullName));

                // Test: basic support to recognize array formulas after reading Excel workbook file
                using (var package = new ExcelPackage(newFile))
                {
                    Assert.AreEqual("A1:A3", package.Workbook.Worksheets["ArrayTest"].Cells["B1"].Formula);
                    Assert.IsTrue(package.Workbook.Worksheets["ArrayTest"].Cells["B1"].IsArrayFormula);
                }
            }
            finally
            {
                File.Delete(newFile.FullName);
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
                //p.SaveAs(new FileInfo(@"c:\temp\issue61.xlsx"));
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
            OpenTemplatePackage("print_titles_170.xlsx");
            _pck.Compatibility.IsWorksheets1Based = false;
            ExcelWorksheet sheet = _pck.Workbook.Worksheets[0];

            sheet.PrinterSettings.RepeatColumns = new ExcelAddress("$A:$C");
            sheet.PrinterSettings.RepeatRows = new ExcelAddress("$1:$3");

            SaveWorkbook("print_titles_170-Saved.xlsx", _pck);
            _pck.Dispose();
        }
        [TestMethod]
        public void Issue172()
        {
            var pck=OpenTemplatePackage("quest.xlsx");
            foreach (var ws in pck.Workbook.Worksheets)
            {
                Console.WriteLine(ws.Name);
            }

            pck.Dispose();
        }

        [TestMethod]
        public void Issue219()
        {
            OpenTemplatePackage("issueFile.xlsx");
            foreach (var ws in _pck.Workbook.Worksheets)
            {
                Console.WriteLine(ws.Name);
            }

            _pck.Dispose();
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
        public void Issue220()
        {
            var pck=OpenPackage("sheetname_pbl.xlsx", true);
            var ws = pck.Workbook.Worksheets.Add("Deal's History");
            var a = ws.Cells["A:B"];
            ws.AutoFilterAddress = ws.Cells["A1:C3"];
            pck.Workbook.Names.Add("Test", ws.Cells["B1:D2"]);
            var name = a.WorkSheet;

            var a2 = new ExcelAddress("'Deal''s History'!a1:a3");
            Assert.AreEqual(a2.WorkSheet, "Deal's History");
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
                Id = Id;
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
            OpenTemplatePackage("Issue236.xlsx");
            _pck.Workbook.Worksheets["Sheet1"].Cells[7, 10].AddComment("test", "Author");
            SaveWorkbook("Issue236-Saved.xlsx", _pck);
        }
        [TestMethod]
        public void Issue228()
        {
            OpenTemplatePackage("Font55.xlsx");
            var ws = _pck.Workbook.Worksheets["Sheet1"];
            var d = ws.Drawings.AddShape("Shape1", eShapeStyle.Diamond);
            ws.Cells["A1"].Value = "tasetraser";
            ws.Cells.AutoFitColumns();
            SaveWorkbook("Font55-Saved.xlsx", _pck);
        }
        [TestMethod]
        public void Issue241()
        {
            var pck=OpenPackage("issue241", true);
            var wks = pck.Workbook.Worksheets.Add("test");
            wks.DefaultRowHeight = 35;
            pck.Save();
            pck.Dispose();
        }
        [TestMethod]
        public void Issue195()
        {
            var pkg = new OfficeOpenXml.ExcelPackage();
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

            pkg.SaveAs(new FileInfo(@"c:\temp\n.xlsx"));
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
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("TextBug");
                ws.Cells["A1"].Value = new DateTime(2019, 3, 7);
                ws.Cells["A1"].Style.Numberformat.Format = "mm-dd-yy";

                Assert.AreEqual("2019-03-07", ws.Cells["A1"].Text);
            }
        }
        [TestMethod]
        public void Issue445()
        {
            ExcelPackage p = new ExcelPackage();
            ExcelWorksheet ws = p.Workbook.Worksheets.Add("AutoFit"); //<-- This line takes forever. The process hangs.
            ws.Cells[1, 1].Value = new string ('a', 50000);
            ws.Cells[1, 1].AutoFitColumns();
}
    }
}
