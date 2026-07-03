using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.RichData;
using OfficeOpenXml.SystemDrawing.Image;
using OfficeOpenXml.SystemDrawing.Text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Threading;
using System.Xml.Linq;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class WorksheetIssues : TestBase
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
        public void s576()
        {
            using (ExcelPackage package = OpenPackage("s576.xlsx", true))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Invoice");

                var namedStyle = package.Workbook.Styles.NamedStyles[0]; // Create a default style
                namedStyle.Style.Font.Name = "Arial";
                namedStyle.Style.Font.Size = 7;

                // Default font and size for spreadsheet  DOES NOT WORK
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 7;

                // Set page size to A4
                worksheet.PrinterSettings.PaperSize = ePaperSize.A4;


                // Set other print settings as needed
                worksheet.PrinterSettings.Orientation = eOrientation.Portrait;
                worksheet.PrinterSettings.FooterMargin = 5;


                // Now 'lines' contains our text split into lines.
                // We can then concatenate these lines with a line break character for the footer.
                //string footerText = string.Join(Environment.NewLine, lines.Take(5)); // Take only the first 5 lines

                var footerText = "This communication is intended only for the addressed recipient(s) and may contain information which is privileged, confidential, commercially sensitive and exempt from " + // + "\n" + 
                    "disclosure under applicable codes and laws.Unauthorised copying.";// or disclosure of this communication to any other person is strictly prohibited. ";// +
                                                                                       //"Please contact the " + //"\n" +
                                                                                       //"undersigned / sender if you are not the intended recipient. "; // + // "\n" +
                                                                                       //																//"MJK Oils Ireland a designated activity company, limited by shares, incorporated in Ireland with registered number 115644 and having its registered office at " + // "\n" +
                                                                                       //																//"Marina Road, Cork, T12 RD92.";


                worksheet.HeaderFooter.OddFooter.LeftAlignedText = footerText;
                worksheet.HeaderFooter.EvenFooter.LeftAlignedText = footerText; // We want the same for even pages

                // Conversion factor (assuming the default font size)
                double conversionFactor = 0.45;


                // Set the widths in millimeters
                worksheet.Column(1).Width = 33 * conversionFactor; // Column A
                worksheet.Column(2).Width = 15 * conversionFactor; // Column B
                worksheet.Column(3).Width = 33 * conversionFactor; // Column C
                worksheet.Column(4).Width = 42 * conversionFactor; // Column D
                worksheet.Column(5).Width = 35 * conversionFactor; // Column E
                worksheet.Column(6).Width = 24 * conversionFactor; // Column F
                worksheet.Column(7).Width = 30 * conversionFactor; // Column G

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s616()
        {
            using (var package = OpenTemplatePackage("s616.xlsx"))
            {
                var Sheet1 = package.Workbook.Worksheets[$"Data Sheet_1"];
                Sheet1.InsertColumn(1, 2);
                var Sheet2 = package.Workbook.Worksheets[$"Data Sheet_2"];
                Sheet2.InsertColumn(1, 2);
                var Sheet3 = package.Workbook.Worksheets[$"Data Sheet_3"];
                Sheet3.InsertColumn(1, 2);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void i1313()
        {
            using (var package = OpenTemplatePackage("SpecialNameValue.xlsx"))
            {
                var sheet = package.Workbook.Worksheets[0];
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void i1314()
        {
            using (var package = OpenTemplatePackage("i1314-2.xlsx"))
            {
                foreach (ExcelWorksheet w in package.Workbook.Worksheets)
                {
                    if (w.Tables.Count() > 0)
                    {
                        var dt = w.Tables.First();
                        if (w == package.Workbook.Worksheets.First()) // First sheet contains the table to be filled by the RAT results
                        {
                            var RowIx = 2;
                            for (int r = 1; r <= 5; r++)
                            {
                                int c = 0;

                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = 1418;
                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = "AfnameNaam";
                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = r;
                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = "VraagNaam";
                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = 1;
                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = 6.2;
                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = "A";
                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = "B";
                                w.Cells[RowIx, dt.Address.Start.Column + c++].Value = 4;
                                var rowRange = dt.AddRow();
                                RowIx = rowRange.Start.Row;
                            }

                            //dt.WorkSheet.Calculate();
                            dt.WorkSheet.Cells.AutoFitColumns();
                            w.Calculate();
                        }

                    }
                }
                package.Save();
                package.Dispose();
            }
        }

        [TestMethod]
        public void AutofitAutofilterTest()
        {
            using var package = OpenPackage("AutofitAutofilterTest.xlsx", true);

            // Two sheets with identical data: one with an autofilter, one without.
            // After autofit, the filtered columns should be wider than the unfiltered
            // ones by the reserved width of the dropdown arrow.
            var wsFilter = package.Workbook.Worksheets.Add("WithFilter");
            var wsNoFilter = package.Workbook.Worksheets.Add("NoFilter");

            foreach (var ws in new[] { wsFilter, wsNoFilter })
            {
                // Headers are the widest text in each column - the data below is deliberately
                // shorter so the column width is driven by the header (+ the dropdown arrow
                // on the filtered sheet).
                ws.Cells["A1"].Value = "Department";
                ws.Cells["B1"].Value = "Annual Budget";
                ws.Cells["C1"].Value = "Region Name";

                ws.Cells["A2"].Value = "Sales";
                ws.Cells["B2"].Value = 1200;
                ws.Cells["C2"].Value = "North";

                ws.Cells["A3"].Value = "IT";
                ws.Cells["B3"].Value = 980;
                ws.Cells["C3"].Value = "West";

                ws.Cells["A4"].Value = "HR";
                ws.Cells["B4"].Value = 540;
                ws.Cells["C4"].Value = "East";
            }

            // Only one sheet gets the autofilter.
            wsFilter.Cells["A1:C4"].AutoFilter = true;

            wsFilter.Cells["A1:C4"].AutoFitColumns();
            wsNoFilter.Cells["A1:C4"].AutoFitColumns();

            for (int col = 1; col <= 3; col++)
            {
                var filterWidth = wsFilter.Column(col).Width;
                var noFilterWidth = wsNoFilter.Column(col).Width;
                System.Diagnostics.Debug.WriteLine($"Column {col}: filter={filterWidth}, noFilter={noFilterWidth}");
                Assert.IsTrue(filterWidth > noFilterWidth,
                    $"Column {col}: filtered width ({filterWidth}) should be greater than unfiltered width ({noFilterWidth}).");
            }

            SaveAndCleanup(package);
        }

        [TestMethod]
        public void i1317()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                package.Workbook.Names.AddValue("ValueName1", 1);
                package.Workbook.Names.AddValue("ValueName2", 2.23);
                package.Workbook.Names.AddValue("ValueName3", true);
                package.Workbook.Names.AddValue("ValueName4", "String Value");
                package.Workbook.Names.AddValue("ValueName5", "String Value with \"");

                package.Save();
                //SaveWorkbook("i1317.xlsx",p);
                using (var p2 = new ExcelPackage(package.Stream))
                {
                    var ws = p2.Workbook.Worksheets[0];
                }
            }
        }
        [TestMethod]
        public void s618()
        {
            using (var package = OpenPackage("s618.xlsx", true))
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet 1");
                var range = worksheet.Cells[2, 1];
                var comment = range.AddComment("Test Comment");
                package.Save();
                worksheet = package.Workbook.Worksheets[0];
                range = worksheet.Cells[2, 1];
                worksheet.Comments.Remove(range.Comment);
                SaveAndCleanup(package);

            }
        }
        [TestMethod]
        public void DeleteRow_TableWithCalculatedColumnFormula()
        {
            using (var pck = new ExcelPackage())
            {
                // Set up a worksheet with a single table that has lots of rows and a calculated column
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                wks.Cells["A1:A14"].Value = "Data outside table";
                wks.Cells["A16"].Value = "Col1";
                wks.Cells["B16"].Value = "Col2";
                var table = wks.Tables.Add(wks.Cells["A16:B18394"], "Table1");
                table.Columns[0].CalculatedColumnFormula = "ROW()-16";

                // The calculated column formula is only given to rows inside the table
                for (int i = 16; i > 0; i--)
                {
                    Assert.AreEqual("", wks.Cells["A" + i].Formula);
                }
                Assert.AreEqual("ROW()-16", wks.Cells["A17"].Formula);

                // Delete all rows in the table except for the header row and the last row
                var listRowsCount = table.Range.Rows;
                wks.DeleteRow(17, listRowsCount - 2);

                // Check that rows above the table haven't been given a formula
                for (int i = 16; i > 0; i--)
                {
                    Assert.AreEqual("", wks.Cells["A" + i].Formula, "Formula present in A" + i);
                }
                Assert.AreEqual("ROW()-16", wks.Cells["A17"].Formula);
                SaveWorkbook("Issue1321.xlsx", pck);
            }
        }
        [TestMethod]
        public void s640()
        {
            using (var package = OpenTemplatePackage("s640.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.First();
                sheet.DeleteRow(6);
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s640_2()
        {
            using (var package = OpenTemplatePackage("s640-2.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.First();
                sheet.DeleteRow(6, 8);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void s641()
        {
            using (var package = OpenTemplatePackage("s641.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.First();
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s668()
        {
            SwitchToCulture("zh");
            try
            {
                using (var package = OpenTemplatePackage("s668.xlsx"))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["test"];
                    try
                    {
                        ExcelCalculationOption excelCalculationOption = new ExcelCalculationOption();
                        excelCalculationOption.AllowCircularReferences = true;
                        worksheet.Calculate(excelCalculationOption);
                    }
                    catch
                    {


                    }
                    SaveAndCleanup(package);
                }
                using (var package = OpenPackage("s668.xlsx"))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["test"];
                    try
                    {
                        ExcelCalculationOption excelCalculationOption = new ExcelCalculationOption();
                        excelCalculationOption.AllowCircularReferences = true;
                        worksheet.Calculate(excelCalculationOption);
                    }
                    catch
                    {


                    }
                    SaveWorkbook("s668-Saved.xlsx", package);
                }
            }
            finally
            {
                SwitchBackToCurrentCulture();
            }

        }
        [TestMethod]
        public void ShareFormulaIDNotFoundError()
        {
            using (var p = OpenTemplatePackage("i1474.xlsx"))
            {
                var ws = p.Workbook.Worksheets.First();
                ws.DeleteRow(35, 2);

                try
                {
                    p.SaveAs("share_formula_error_test.xlsx");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }
        [TestMethod]
        public void s720()
        {
            using (var p = OpenTemplatePackage("s720.xlsx"))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets[0];

                try
                {
                    worksheet.Cells["A1:A3"].Insert(eShiftTypeInsert.Right);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"error {ex}");
                }

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s721()
        {
            using (var p = OpenTemplatePackage("s721.xlsx"))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets["sheet1"];
                Assert.AreEqual(ePhoneticType.NoConversion, worksheet.PhoneticProperties.PhoneticType);
                Assert.AreEqual(ePhoneticAlignment.Left, worksheet.PhoneticProperties.Alignment);
                Assert.AreEqual(1, worksheet.PhoneticProperties.FontId);

                var formulaD2 = p.Workbook.Worksheets["Sheet2"].Cells["D2"].Formula;
                p.Save();

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    Assert.AreEqual(formulaD2, p2.Workbook.Worksheets["Sheet2"].Cells["D2"].Formula);
                }
            }
        }
        [TestMethod]
        public void DimensionValueIssue()
        {
            using (var excelPackage = OpenTemplatePackage(@"s719-DimensionByValue.xlsx"))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["1"];

                Console.WriteLine(excelWorksheet.Dimension.Columns);
                Console.WriteLine(excelWorksheet.DimensionByValue.Columns);
            }
        }
        [TestMethod]
        public void s730()
        {
            using (var p = OpenTemplatePackage("s730.xlsx"))
            {
                string sheetName = "披露附注";
                var ws = p.Workbook.Worksheets[sheetName];
                ws.Cells["G8700:G8705"].Insert(eShiftTypeInsert.Right);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateShiftRightSecondPage_CellStore()
        {
            using (var p = OpenPackage("s730-2.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.SetValue(8244, 7, "x");
                ws.Cells["G8700:G8707"].Style.Fill.SetBackground(Color.Yellow, OfficeOpenXml.Style.ExcelFillStyle.Solid);
                ws.Cells["G8700:G8705"].Insert(eShiftTypeInsert.Right);

                Assert.AreEqual("x", ws.GetValue(8244, 7));
                Assert.AreEqual("FFFFFF00", ws.Cells["H8700"].Style.Fill.BackgroundColor.Rgb);
                Assert.AreEqual("FFFFFF00", ws.Cells["H8705"].Style.Fill.BackgroundColor.Rgb);
                Assert.IsNull(ws.Cells["H8706"].Style.Fill.BackgroundColor.Rgb);
                Assert.IsNull(ws.Cells["H8707"].Style.Fill.BackgroundColor.Rgb);

                Assert.AreEqual("FFFFFF00", ws.Cells["G8706"].Style.Fill.BackgroundColor.Rgb);
                Assert.AreEqual("FFFFFF00", ws.Cells["G8707"].Style.Fill.BackgroundColor.Rgb);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void I1596()
        {
            using (var p = OpenTemplatePackage("i1596.xlsx"))
            {
                ExcelWorkbook workbook = p.Workbook;
                ExcelWorksheet worksheet = workbook.Worksheets[1];

                worksheet.DeleteRow(256);
            }
        }
        [TestMethod]
        public void s746()
        {
            using (var p = OpenTemplatePackage("s746.xlsm"))
            {
                var workbook = p.Workbook;
                var worksheet = workbook.Worksheets["Sheet1"];
                workbook.Worksheets["Sheet1"].Columns[2].Width = 100; //Commenting this line out stops the error.
                SaveAndCleanup(p);

            }
        }
        [TestMethod]
        public void i1663()
        {
            using (var p1 = OpenTemplatePackage("i1663-source.xlsx"))
            {
                var copiedSht = p1.Workbook.Worksheets[0];
                using (var p2 = OpenTemplatePackage("i1663-dest.xlsx"))
                {
                    p2.Workbook.Worksheets.Add("newSht", copiedSht);
                    SaveAndCleanup(p2);
                }
            }
        }
        [TestMethod]
        public void I1628()
        {
            using (var p = OpenPackage("i1628.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "A\r\n\tB";
                SaveAndCleanup(p);

            }
        }
        [TestMethod]
        public void I1691()
        {
            using (var p = OpenTemplatePackage("i1691.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void I1728()
        {
            using var p = OpenTemplatePackage("Issue1728.xlsm");
            var nWs = p.Workbook.Worksheets.Count;
            var i = 0;
            foreach (var ws in p.Workbook.Worksheets)
            {
                i++;
                var dimensionRows = ws.Dimension.Rows;
                var dimensionByValueRows = ws.DimensionByValue.Rows;
            }
        }

        [TestMethod]
        public void i1742()
        {
            // before this fix we couldn't delete the very last column on the sheet...
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var maxCol = ExcelPackage.MaxColumns;
            sheet.DeleteColumn(maxCol);
        }
        [TestMethod]
        public void i1709()
        {
            using (var p = OpenTemplatePackage("i1709.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void i1709_2()
        {
            using (var p = OpenPackage("i1709-2.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "row 1_x000d__x000d_col 1";
                ws.Cells["A2"].Value = "row 2\r\rcol 1";
                ws.Cells["A3"].Value = "row 3\r\n\r\ncol 1";
                ws.Cells["A4"].Value = "row 4\n\ncol 1";
                ws.Cells["A5"].Value = "row 5_x000d_\ncol 1";
                ws.Cells["A6"].Value = "row 6_x000d__x000a_col 1";

                ws.Cells["A1:A6"].Style.WrapText = true;
                ws.Cells["B1:B6"].Formula = "=CODE(MID(A1,5,1))";
                ws.Cells["C1:C6"].Formula = "=CODE(MID(A1,6,1))";
                ws.Cells["D1:D6"].Formula = "=CODE(MID(A1,7,1))";
                ws.Cells["E1:E6"].Formula = "=CODE(MID(A1,8,1))";
                SaveAndCleanup(p);
            }
        }
        private class I1782DataItem
        {
            public int Id { get; set; }
            [DisplayName("Project Number")]
            public ExcelHyperLink ProjectNumberUrl
            {
                get;
                set;
            }
        }
        [TestMethod]
        public void i1782()
        {
            var list = new List<I1782DataItem>();
            var hl = new ExcelHyperLink("https://epplussoftware.com", "epplussoftware.com");
            list.Add(new I1782DataItem { Id = 1, ProjectNumberUrl = hl });

            using var p = OpenPackage("i1782.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("sheet1");
            ws.Cells["A1"].LoadFromCollection(list, true, OfficeOpenXml.Table.TableStyles.None, BindingFlags.Instance | BindingFlags.Public, new[] { typeof(I1782DataItem).GetProperty("Id"), typeof(I1782DataItem).GetProperty("ProjectNumberUrl") });

            Assert.IsNotNull(ws.Cells["B2"].Hyperlink);

            SaveAndCleanup(p);
        }
        [TestMethod]
        public void s787()
        {
            using var p = OpenPackage("s787.xlsx", true);

            var renamedWorksheet = p.Workbook.Worksheets.Add("RenamedWorksheet");
            renamedWorksheet.Cells[1, 1].Value = "Value";

            var referencingWorksheet = p.Workbook.Worksheets.Add("ReferencingWorksheet");
            referencingWorksheet.Cells[1, 1].Formula = "=RenamedWorksheet!A1";

            renamedWorksheet.Name = "Renamed Worksheet";
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void Issue1794_1()
        {
            // This tests creates a workbook without errors. When this workbook is opened in Excel
            // and then closed without changing anything, Excel still shows a "Save changes" dialog.
            // this seems to be related to that Excel renames the worksheet xml files.
            // EPPlus keeps the sheet2.xml and sheet3.xml file names after the line p.Workbook.Worksheets.Delete(wsTemplate);
            // this bug was fixed in GitHub Issue 1794 /MA

            using var p = OpenTemplatePackage("Issue1794.xltx");
            var wsTemplate = p.Workbook.Worksheets[0];

            for (int i = 0; i < 2; i++)
            {
                var ws = p.Workbook.Worksheets.Add(i.ToString(), wsTemplate);
                ws.View.SetTabSelected();      // avoids grouping
            }
            p.Workbook.Worksheets.Delete(wsTemplate);
            SaveWorkbook("Issue1794_1_Output.xlsx", p);
        }

        [TestMethod]
        public void Issue1794_2()
        {
            // This tests creates a workbook without errors. When this workbook is opened in Excel
            // and then closed without changing anything, Excel still shows a "Save changes" dialog.
            // this seems to be related to that Excel renames the worksheet xml files.
            // EPPlus keeps the sheet2.xml and sheet3.xml file names after the line p.Workbook.Worksheets.Delete(wsTemplate);
            // this bug was fixed in github Issue 1794 /MA
            using var p = OpenTemplatePackage("Issue1794.xlsx");
            var wsTemplate = p.Workbook.Worksheets[0];

            for (int i = 0; i < 2; i++)
            {
                var ws = p.Workbook.Worksheets.Add(i.ToString(), wsTemplate);
                ws.View.SetTabSelected();      // avoids grouping
            }
            wsTemplate.View.SetTabSelected(false);
            p.Workbook.Worksheets.Delete(wsTemplate);
            SaveWorkbook("Issue1794_2_Output.xlsx", p);
        }
        [TestMethod]
        public void DeletingWorksheetsWithParameters()
        {
            using (var p = OpenPackage("DeletingGroupOfWorksheets.xlsx", true))
            {
                var wb = p.Workbook;
                var worksheets = wb.Worksheets;

                for (int i = 0; i < 5; i++)
                {
                    worksheets.Add($"Data {i}");
                }

                for (int i = 0; i < 5; i++)
                {
                    worksheets.Add($"SomeWorksheet{i}");
                }

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    var ws = p.Workbook.Worksheets[i];
                    if (ws.Name.StartsWith("Data ", StringComparison.OrdinalIgnoreCase))
                    {
                        p.Workbook.Worksheets.Delete(ws);
                        i--;
                    }
                }
                var countWs = p.Workbook.Worksheets.Count;

                Assert.AreEqual(countWs, 5);

                worksheets.Delete($"SomeWorksheet2");

                Assert.AreEqual(p.Workbook.Worksheets.Count, 4);
                Assert.AreEqual("SomeWorksheet0", p.Workbook.Worksheets[0].Name);
                Assert.AreEqual("SomeWorksheet1", p.Workbook.Worksheets[1].Name);
                Assert.AreEqual("SomeWorksheet3", p.Workbook.Worksheets[2].Name);
                Assert.AreEqual("SomeWorksheet4", p.Workbook.Worksheets[3].Name);

                Assert.AreEqual("SomeWorksheet0", p.Workbook.Worksheets["SomeWorksheet0"].Name);
                Assert.AreEqual("SomeWorksheet1", p.Workbook.Worksheets["SomeWorksheet1"].Name);
                Assert.AreEqual("SomeWorksheet3", p.Workbook.Worksheets["SomeWorksheet3"].Name);
                Assert.AreEqual("SomeWorksheet4", p.Workbook.Worksheets["SomeWorksheet4"].Name);

            }
        }
        [TestMethod]
        public void DeletingWorksheetsWithParameters_1Base()
        {
            using (var p = OpenPackage("DeletingGroupOfWorksheets.xlsx", true))
            {

                p.Compatibility.IsWorksheets1Based = true;
                var wb = p.Workbook;
                var worksheets = wb.Worksheets;

                for (int i = 1; i <= 5; i++)
                {
                    worksheets.Add($"Data {i}");
                }

                for (int i = 1; i <= 5; i++)
                {
                    worksheets.Add($"SomeWorksheet{i}");
                }

                for (int i = 1; i <= p.Workbook.Worksheets.Count; i++)
                {
                    var ws = p.Workbook.Worksheets[i];
                    if (ws.Name.StartsWith("Data ", StringComparison.OrdinalIgnoreCase))
                    {
                        p.Workbook.Worksheets.Delete(ws);
                        i--;
                    }
                }
                var countWs = p.Workbook.Worksheets.Count;

                Assert.AreEqual(countWs, 5);

                worksheets.Delete($"SomeWorksheet2");

                Assert.AreEqual(p.Workbook.Worksheets.Count, 4);
                Assert.AreEqual("SomeWorksheet1", p.Workbook.Worksheets[1].Name);
                Assert.AreEqual("SomeWorksheet3", p.Workbook.Worksheets[2].Name);
                Assert.AreEqual("SomeWorksheet4", p.Workbook.Worksheets[3].Name);
                Assert.AreEqual("SomeWorksheet5", p.Workbook.Worksheets[4].Name);

                Assert.AreEqual("SomeWorksheet1", p.Workbook.Worksheets["SomeWorksheet1"].Name);
                Assert.AreEqual("SomeWorksheet3", p.Workbook.Worksheets["SomeWorksheet3"].Name);
                Assert.AreEqual("SomeWorksheet4", p.Workbook.Worksheets["SomeWorksheet4"].Name);
                Assert.AreEqual("SomeWorksheet5", p.Workbook.Worksheets["SomeWorksheet5"].Name);

            }
        }
        [TestMethod]
        public void s816()
        {
            using var excelPackage = OpenTemplatePackage("s816.xlsx");
            var sheet = excelPackage.Workbook.Worksheets.First();

            // Act
            sheet.Cells.Sort(column: 0);

            var commentText = sheet.Cells["A3"].Comment.Text;
            Assert.AreEqual("6", commentText);

            excelPackage.Save();

            using var loadedExcelPackage = new ExcelPackage(excelPackage.Stream);
            var loadedSheet = loadedExcelPackage.Workbook.Worksheets.First();

            var loadedCommentText = loadedSheet.Cells["A3"].Comment.Text;
            Assert.AreEqual("6", loadedCommentText);
        }
        [TestMethod]
        public void properties()
        {
            using (var package = OpenTemplatePackage("properties.xlsx"))
            {
                package.Workbook.Properties.LastModifiedBy = "";
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void s816_2()
        {
            using var excelPackage = OpenTemplatePackage("s816-2.xlsx");
            var sheet = excelPackage.Workbook.Worksheets.First();
            var formula = sheet.Cells["B5"].Formula;
            // Act
            sheet.Cells.Sort(column: 0);

            Assert.AreEqual(sheet.Cells["B3"].Formula, formula);
            SaveAndCleanup(excelPackage);
        }
        [TestMethod]
        public void i1870()
        {
            using var savedExcelPackage = OpenTemplatePackage("i1870.xlsx");
            var sheet = savedExcelPackage.Workbook.Worksheets.First();

            // Act
            sheet.Cells["2:3"].Clear();
            sheet.Cells["6:6"].Clear();
            sheet.Cells.Sort(column: 0);

            //Assert 1

            Assert.AreEqual("2", sheet.Cells["A1"].ThreadedComment.Comments.First().Text);
            Assert.AreEqual("3", sheet.Cells["A2"].ThreadedComment.Comments.First().Text);
            Assert.AreEqual("2", sheet.Cells["B1"].Comment.Text);
            Assert.AreEqual("3", sheet.Cells["B2"].Comment.Text);

            //Act 2
            SaveWorkbook("i1870-save.xlsx", savedExcelPackage);

            //Assert 2
            using var loadedExcelPackage = OpenPackage("i1870-save.xlsx");
            var loadedSheet = loadedExcelPackage.Workbook.Worksheets.First();

            Assert.AreEqual("2", loadedSheet.Cells["A1"].ThreadedComment.Comments.First().Text);
            Assert.AreEqual("3", loadedSheet.Cells["A2"].ThreadedComment.Comments.First().Text);
            Assert.AreEqual("2", loadedSheet.Cells["B1"].Comment.Text);
            Assert.AreEqual("3", loadedSheet.Cells["B2"].Comment.Text);
        }
        [TestMethod]
        public void i1876()
        {
            using (var p = OpenTemplatePackage("i1876.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var dv = ws.DimensionByValue;

                Assert.AreEqual("A1:F1", dv.Address);

            }
        }
        [TestMethod]
        public void i1878()
        {
            using (var p = OpenTemplatePackage("i1878.xlsx"))
            {
                var ws = p.Workbook.Worksheets.First();

                var timeSpanCell = ws.GetValue<TimeSpan>(1, 1);

                Assert.AreEqual(timeSpanCell.Ticks, new TimeSpan(12, 30, 45).Ticks);
            }
        }
        [TestMethod]
        public void s843()
        {
            using var excelPackage = OpenTemplatePackage("s843.xlsx");
            var sheet = excelPackage.Workbook.Worksheets.First();

            sheet.Cells.Sort(0);

            var existingThread = sheet.ThreadedComments.Threads.Single();
            sheet.ThreadedComments.Remove(existingThread);

            sheet.Cells["C3"].AddComment("Test"); // NullReferenceException 
        }
        [TestMethod]
        public void i1951()
        {
            using (var p = OpenPackage("I1951.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("GenericTM");

                AddMeasureSheet(p, ws);

                p.Settings.TextSettings.PrimaryTextMeasurer = new SystemDrawingTextMeasurer();
                ws = p.Workbook.Worksheets.Add("SystemDrawingTM");
                AddMeasureSheet(p, ws);

                SaveAndCleanup(p);
            }
        }

        private static void AddMeasureSheet(ExcelPackage p, ExcelWorksheet ws)
        {
            string multiLineText = "Line one" + Environment.NewLine + "Line two is longer" + "\n" + "Extra line";

            ws.Cells["A1"].Value = multiLineText;
            ws.Cells["A1"].Style.WrapText = true;

            ws.Cells["B2"].Value = multiLineText;

            ws.Cells["C1"].Value = multiLineText;
            ws.Cells["A1"].Style.WrapText = true;

            ws.Cells["B2"].Value = multiLineText;

            p.Settings.TextSettings.MeasureWrappedTextCells = true;
            // AutoFitColumns - calculates width as if there were no line breaks.
            ws.Cells["A1:B2"].AutoFitColumns();

            p.Settings.TextSettings.MeasureWrappedTextCells = false;
            ws.Cells["C1"].Value = multiLineText;
            ws.Cells["C1"].Style.WrapText = true;
            ws.Cells["D2"].Value = multiLineText;

            // AutoFitColumns - calculates width as if there were no line breaks.
            ws.Cells["C1:D2"].AutoFitColumns();
        }

        //i2084
        [TestMethod]
        public void s912_Alternate()
        {
            //Optimizing for not overwriting existing styles
            using (var package = OpenPackage("s912_alt.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("F1");

                int nbLines = 10000;
                int nbCols = 100;

                var sw = new Stopwatch();
                sw.Start();

                for (int i = 1; i <= nbLines; i++)
                {
                    for (int j = 1; j <= nbCols; j++)
                    {
                        var cell = sheet.Cells[i, j];
                        var cellNumberFormat = cell.Style.Numberformat;
                        cell.Value = 123;
                    }
                }
                sw.Stop();

                var seconds = sw.Elapsed.TotalSeconds;
                Assert.IsTrue(seconds < 10.0D);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void s912()
        {
            using (var package = OpenPackage("s912.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("F1");

                int nbLines = 10000;
                int nbCols = 100;

                var sw = new Stopwatch();
                sw.Start();

                // Uncommenting one of these lines changes the performance of the for loops.
                // At the end of each line is the measured time of the whole program, when this
                // specific line is uncommented. When no line is uncommented, the measured time
                // is 12.7s.
                //
                // sheet.Cells[1, 1, nbLines, nbCols].Style.Numberformat.Format = "#"; // 7.4s
                // sheet.Cells[1, 1, nbLines, nbCols].Style.Locked = true; // 7.4s
                //sheet.Cells[1, 1, nbLines, nbCols].Value = 1; // 19.5s // uncommenting this alone is ~2s after fix. With below about 3s. 
                // sheet.Cells[1, 1, nbLines, nbCols].Value = ""; // 18s
                // sheet.InsertColumn(1, nbCols); // 12.9
                // sheet.InsertColumn(1, nbCols, 1); // 7.8s

                for (int i = 1; i <= nbLines; i++)
                {
                    for (int j = 1; j <= nbCols; j++)
                    {
                        sheet.SetValue(i, j, 123);
                    }
                }

                var seconds = sw.Elapsed.TotalSeconds;
                sw.Stop();

                //seconds was ~1.5-1.7 locally in 8.0.9
                //Made test check if over 10 seconds in case of slow appveyor
                Assert.IsTrue(seconds < 10.0D);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void DimensionByValueIssue()
        {
            using (var p = OpenTemplatePackage("DimensionByValueError.xlsx"))
            {
                var ws = p.Workbook.Worksheets["Technical"];
                var dv = ws.DimensionByValue;
                Assert.AreEqual("C3", dv.Start.Address);
                Assert.AreEqual("M60", dv.End.Address);
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void GermanCultureFormattingResultsInError()
        {
            var excelPackage = OpenTemplatePackage("Test_TextFormular.xlsx");
            excelPackage.Workbook.NumberFormatToTextHandler = options => "64.066,27€";
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("de-DE");
            excelPackage.Workbook.Calculate();
            Assert.AreEqual("64.066,27€", excelPackage.Workbook.Worksheets[0].Cells[1, 3].Text);
        }

        [TestMethod]
        public void s931()
        {
            using (ExcelPackage xlPackage = OpenPackage("s931.xlsx", true))
            {
                ExcelWorksheet sheet = xlPackage.Workbook.Worksheets.Add("test");

                sheet.Cells[1, 1].Value = "a";
                sheet.Cells[2, 1].Value = "b";
                sheet.Cells[3, 1].Value = "c";
                sheet.Cells[4, 1].Value = "d";
                sheet.Cells[5, 1].Value = "e";
                sheet.Cells[6, 1].Value = "f";
                sheet.Cells[7, 1].Value = "g";
                sheet.Cells[8, 1].Value = "h";
                sheet.Cells[9, 1].Value = "i";
                sheet.Cells[10, 1].Value = "j";
                sheet.Cells[11, 1].Value = "k";
                sheet.Cells[12, 1].Value = "l";

                sheet.Row(1).Hidden = false;
                sheet.Row(2).Hidden = false;

                sheet.Row(3).Hidden = true;
                sheet.Row(4).Hidden = true;
                sheet.Row(5).Hidden = true;
                sheet.Row(6).Hidden = true;
                sheet.Row(7).Hidden = true;
                sheet.Row(8).Hidden = true;
                sheet.Row(9).Hidden = true;
                sheet.Row(10).Hidden = true;

                sheet.View.FreezePanes(11, 1);

                sheet.Row(6).Hidden = false;
                sheet.Row(7).Hidden = false;

                //sheet.View.PaneSettings.YSplit = 10;

                Assert.AreEqual(10, sheet.View.PaneSettings.YSplit);

                ExcelWorksheet sheet2 = xlPackage.Workbook.Worksheets.Add("test2");

                sheet2.Column(1).Hidden = true;
                sheet2.Row(1).Hidden = true;
                var address = sheet2.Cells["C3"];

                var row = address.Start.Row;
                var col = address.Start.Column;

                address.Value = "Freeze here";
                sheet2.View.FreezePanes(address.Start.Row, address.Start.Column);

                var someval = sheet2.View.PaneSettings.YSplit;

                //Assert.AreEqual(2, sheet2.View.PaneSettings.YSplit);

                xlPackage.Save();
            }
        }
        [TestMethod]

        public void Issue2157()
        {
            using (var p = OpenTemplatePackage("i2157.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var dv = ws.DimensionByValue;
                Assert.AreEqual(2, dv.Columns);
            }
        }
        [TestMethod]
        public void s975()
        {
            using (ExcelPackage package = OpenTemplatePackage("s975.xlsx"))
            {
                var ws = package.Workbook.Worksheets["抵消附注"];
                var cell = ws.Cells["E2419"];
                cell.Value = null;
            }
        }
        [TestMethod]
        public void issue2191()
        {
            using var pkg = new ExcelPackage();
            var ws = pkg.Workbook.Worksheets.Add("Sheet1");

            // 1. Style a cell in column A (without value)
            ws.Cells["A1"].Style.Font.Bold = true;

            // 2. Apply row-level style (triggers customFormat in XML)
            ws.Row(2).Style.Font.Name = "Arial";

            // 3. Set value in column B (not A) on the styled row
            ws.Cells["B2"].Value = "Test";

            // Save and reload
            var ms = new MemoryStream();
            pkg.SaveAs(ms);
            ms.Position = 0;

            using var pkg2 = new ExcelPackage(ms);
            var ws2 = pkg2.Workbook.Worksheets[0];

            // This throws "Column out of range"
            var dbv = ws2.DimensionByValue;

            Assert.AreEqual("B2", dbv.Address);
        }
        [TestMethod]
        public void i2240()
        {
            using (var package = OpenTemplatePackage("i2240.xlsx"))
            {
                var theText = package.Workbook.Worksheets.First().HeaderFooter.OddFooter.LeftAlignedText;

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void testRepro()
        {
            SwitchToCulture("");
            using (ExcelPackage p = OpenTemplatePackage("reproTimesheets - Copy.xlsx"))
            {
                var Styles = p.Workbook.Styles;

                p.Workbook.Worksheets[0].Cells["A10"].Style.Font.Bold = true;
                p.Workbook.Worksheets[0].Cells["A10"].Value = "Debugging";

                p.Workbook.Worksheets.Add("SomeSheet");
                p.Workbook.Worksheets[1].Cells["A10"].Value = "Debugging2";

                Stream fs = File.Create(("C:\\epplusTest\\Testoutput\\" + "reproTimesheets.xlsx").Replace("file://", ""));
                p.SaveAs(fs);
                fs.Close();
            }
            SwitchBackToCurrentCulture();
        }

        [TestMethod]
        public void i2258()
        {
            var p = OpenTemplatePackage("repro2.xlsx");
            var ws = p.Workbook.Worksheets.First();

            var a1 = ws.Cells["A1"].Text;
            Assert.AreEqual("-/- 1 000", a1);

            var b1 = ws.Cells["B1"].Text;
            Assert.AreEqual("-/- 2 000", b1);

            var c1 = ws.Cells["C1"].Text;
            Assert.AreEqual("3 000,00", c1);

            var d1 = ws.Cells["D1"].Text;
            Assert.AreEqual("(4000,000)", d1);

            var a3 = ws.Cells["A3"].Text;
            Assert.AreEqual("1000,000", a3);

            var b3 = ws.Cells["B3"].Text;
            Assert.AreEqual("(2000,000)", b3);
            
            var c3 = ws.Cells["C3"].Text;
            Assert.AreEqual("-3000,000", c3);
            
            var d3 = ws.Cells["D3"].Text;
            Assert.AreEqual("Negative 4000,000", d3);
        }

        [TestMethod]
        public void InsertAndShift_ShouldNotThrow_WhenArrayIsFull()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            // Fill 7 columns with formatting to trigger the boundary condition
            for (int col = 1; col <= 7; col++)
            {
                sheet.Column(col).Width = 15;
            }

            // These two inserts should not throw ArgumentException
            sheet.InsertColumn(3, 1);
            sheet.InsertColumn(5, 1);
        }

        [TestMethod]
        public void Issue2325()
        {
            var ex = Assert.ThrowsExactly<NotSupportedException>(() =>
            {
                var package = OpenTemplatePackage("StrictOpenXml.xlsx");
                var ws = package.Workbook.Worksheets.First();
            });
            Assert.IsTrue(ex.Message.Contains("Strict Open XML"));
        }

    }
}
