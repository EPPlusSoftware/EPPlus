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
  04/03/2025         EPPlus Software AB       Initial release EPPlus 8.1
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlusTest.LongRunning
{
    [TestClass, Ignore]
    public class LongRunningIssuesTests : TestBase
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
        public void s350()
        {
            using (var p = OpenTemplatePackage("s350.xlsm"))
            {
                SaveWorkbook("s350.xlsm", p);
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
        public void s551_2()
        {
            using (var p = OpenTemplatePackage("s551.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var usedRange = ws.Cells["a1:b5"];
                foreach (ExcelRangeRow dataRow in usedRange.EntireRow)
                {
                    if (dataRow.Hidden == false)
                    {
                        dataRow.Range.Formula = "f1";
                    }
                }
            }
        }
        [TestMethod]
        public void i863()
        {
            using (var p = OpenTemplatePackage("i863.xlsx"))
            {
                // Removed insertion of PHI data, just re-saving the template for sample purposes

                // Workaround - Issue with "Inputs" tab - Validation of T60:T64 failed: Formula2 must be set if operator is 'between' or 'notBetween' when cells are not using between or notBetween
                var otherInputTab = p.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals("Inputs"));
                if (otherInputTab != null)
                {
                    otherInputTab.DataValidations.InternalValidationEnabled = false;
                }
                // Saving
                SaveAndCleanup(p);

                var p2 = OpenPackage("i863.xlsx");

                var ws17 = p2.Workbook.Worksheets[16];
            }
        }
        [TestMethod]
        public void s539()
        {
            //Outputs
            var pc = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

                string sheetName = "Sheet1";
                string range = "G2:G5";
                string value = "VLOOKUP(F2,'Reference Data'!A2:B187021,2,0)";
                var logFile = new FileInfo("c:\\temp\\formulaLog.log");
                if (logFile.Exists) logFile.Delete();
                using (var package = OpenTemplatePackage("s539.xlsm"))
                {
                    package.Workbook.FormulaParserManager.AttachLogger(logFile);
                    var ws = package.Workbook.Worksheets[sheetName];
                    ws.Cells[range].Formula = value;
                    ws.Cells[range].Calculate();
                    SaveAndCleanup(package);
                }
            }
            catch (Exception e)
            {
                string exc = "";
                exc = "Failed. " + e.ToString();
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = pc;
                System.GC.Collect();
            }
        }
        [TestMethod]
        public void s610()
        {
            using (var p = OpenTemplatePackage("s610.xlsx"))
            {
                var wTestSheet = p.Workbook.Worksheets[0];
                //wTestSheet.Name = "Sheet2";
                //wTestSheet.View.UnFreezePanes();
                wTestSheet.InsertColumn(1, 2);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s614()
        {
            using (var package = OpenTemplatePackage("s614.xlsx"))
            {
                int sheetIndex = 5;
                var sheetName = $"Data Sheet_{sheetIndex}";
                var worksheet = package.Workbook.Worksheets[sheetName];
                worksheet.Name = "TestSheet_{sheetIndex}";

                worksheet.InsertColumn(1, 2);
                worksheet.Cells.Style.Font.Name = "ＭＳ Ｐゴシック";
                worksheet.Cells.Style.Font.Size = 11;

                worksheet.Cells[1, 1].Value = "TextTextTextTextTextTextTextTextTextTextTextText";

                worksheet.Column(1).AutoFit();
                worksheet.Column(2).AutoFit();

                package.Save();
            }
        }
        [TestMethod]
        public void s789()
        {
            using (var package = OpenTemplatePackage("s789.xlsx"))
            {
                var wb = package.Workbook;
                foreach (var ws in package.Workbook.Worksheets)
                {
                    foreach (var pTable in ws.PivotTables)
                    {
                        foreach (var field in pTable.Fields)
                        {
                        }
                    }
                }

                SaveAndCleanup(package);
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
        public void SaveDefinedName()
        {
            using (var p = OpenTemplatePackage("SaveIssueName.xlsm"))
            {
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
        public void s831()
        {
            using var p = OpenTemplatePackage("s831.xlsx");
            var sheet = p.Workbook.Worksheets[0];
            var sw = new Stopwatch();
            sw.Start();
            p.Workbook.Calculate();
            //p.Workbook.FormulaParser.
            GC.Collect();

            Console.WriteLine(new DateTime(sw.ElapsedTicks).ToString("HH:mm:ss"));
        }
        [TestMethod]
        public void s688()
        {
            using (ExcelPackage package = OpenTemplatePackage("s688.xlsx"))
            {
                package.Workbook.Worksheets[0].PivotTables[0].Calculate(false);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void s463()
        {
            using (var p = OpenTemplatePackage("SRK2016.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void s569()
        {
            var sheetName = "披露表(国资)";

            using (var p = OpenTemplatePackage("s569source.xlsx"))
            {
                var SourceWB = p.Workbook;
                using (var tP = OpenTemplatePackage("s569target.xlsm"))
                {
                    var tBook = tP.Workbook;
                    var sSheet = p.Workbook.Worksheets.GetByName(sheetName);
                    tBook.Worksheets.Add(sheetName, sSheet);

                    SaveAndCleanup(tP);
                }
            }
        }

        #region ConditionalFormatting Issues
        [TestMethod]
        public void s725()
        {
            using (var p1 = OpenTemplatePackage("s725.xlsx"))
            {
                var sheet = p1.Workbook.Worksheets[6];
                if (p1.Workbook.Worksheets.Count > 0)
                {
                    p1.Save();
                }
                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    var sheet2 = p2.Workbook.Worksheets[6];
                    SaveWorkbook("s725-secondsaveorig.xlsx", p2);
                }
            }
        }
        [TestMethod]
        public void s782()
        {
            using (var package = OpenTemplatePackage("s782.xlsx"))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["披露附注"];

                string areaStr = "E247:E256";
                worksheet.Cells[areaStr].Insert(eShiftTypeInsert.Right);

                SaveAndCleanup(package);
            }
        }
        #endregion
        #region PivotTableIssues
        [TestMethod]
        public void s744()
        {
            using (var p = OpenTemplatePackage("s744.xlsx"))
            {
                ExcelWorkbook workbook = p.Workbook;
                SaveAndCleanup(p);
            }
        }
        #endregion
        #region DefinedNameIssues
        [TestMethod]
        public void I1238()
        {
            using (var p = OpenTemplatePackage("I1238SlowWorkbook.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.Cells["A1"].Value = 1;
                SaveAndCleanup(p);
            }
        }
        #endregion
        #region FormulaCalculationIssues
        [TestMethod]
        public void i1540()
        {
            using (var p = OpenPackage("i1540.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "A";
                ws.Cells["A2"].Value = "B";
                ws.Cells["A3"].Value = "C";
                ws.Cells["B1:B3"].FillNumber(1, 1);
                ws.Cells["C1:C3"].FillNumber(10, 10);
                ws.Cells["E1"].Formula = "SUM(If(A:A=\"A\",B:B,C:C))";                          //Should be set as an array formula
                ws.Cells["E2"].Formula = "SUM(If(A1:A3=\"A\",B1:B3,C1:C3))";                    //Should be set as an array formula
                ws.Cells["F1"].Formula = "SUM(If(A:A=\"A\",B:B,C:C))";                          //Should be set as an array formula
                ws.Cells["F2"].Formula = "SUM(If(A1:A3=\"A\",B1:B3,C1:C3))";                    //Should be set as an array formula
                ws.Cells["F1:F2"].UseImplicitItersection = true;

                ws.Cells["G1"].CreateArrayFormula("SUM(If(A:A=\"A\",B:B,C:C))", true);
                ws.Cells["G2"].CreateArrayFormula("SUM(If(A1:A3=\"A\",B1:B3,C1:C3))", true);

                ws.Cells["E1:G2"].Calculate();

                Assert.AreEqual(51D, ws.Cells["E1"].Value); //Will be handled as a dynamic formula when calculated, not as in Excel where implicit intersections seems to be applied inside the sum.
                Assert.AreEqual(51D, ws.Cells["E2"].Value);
                Assert.AreEqual(6D, ws.Cells["F1"].Value);
                Assert.AreEqual(60D, ws.Cells["F2"].Value);

                SaveAndCleanup(p);
            }
        }
        #endregion
        #region WorksheetIssues
        [TestMethod]
        public void s775()
        {
            string sheetName = "披露附注";

            List<int> add = new List<int>()
            {
                4,9,15
            };
            using (ExcelPackage package = OpenTemplatePackage("s775.xlsx"))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                ExcelNamedRange namedRange = worksheet.Names["_jds1165020120230"];
                int startRow = namedRange.Start.Row;

                var cell = worksheet.Cells["D2059"];
                var cell2 = worksheet.Cells["D2060"];

                worksheet.InsertRow(2059, 1, 2059 - 1);

                package.Save();
            }
        }
        #endregion
    }
}
