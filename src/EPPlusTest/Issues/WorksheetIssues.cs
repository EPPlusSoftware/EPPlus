using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System.Globalization;
using System.Threading;
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

				//var namedStyle = package.Workbook.Styles.CreateNamedStyle("Default"); // Create a default style
				//namedStyle.Style.Font.Name = "Arial";
				//namedStyle.Style.Font.Size = 7;
				var namedStyle = package.Workbook.Styles.NamedStyles[0]; // Create a default style
				namedStyle.Style.Font.Name = "Arial";
				namedStyle.Style.Font.Size = 7;

				//"&L&\"Arial,Normal\"&8";


				// Default font and size for spreadsheet  DOES NOT WORK
				worksheet.Cells.Style.Font.Name = "Arial";
				worksheet.Cells.Style.Font.Size = 7;

				// Set page size to A4
				worksheet.PrinterSettings.PaperSize = ePaperSize.A4;


				// Set other print settings as needed
				worksheet.PrinterSettings.Orientation = eOrientation.Portrait;
				//worksheet.PrinterSettings.FitToPage = true;
				//worksheet.PrinterSettings.FitToWidth = 1;
				worksheet.PrinterSettings.FooterMargin = 5;




				//string longText = ""; // Our long string
				//int maxLineLength = 140; // Maximum length of each line, adjust as needed
				//						 //var lines = SplitStringIntoLines(longText, maxLineLength);


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


				//worksheet.HeaderFooter.OddFooter.CenteredText = "Test Disclaimer";
				//worksheet.HeaderFooter.EvenFooter.CenteredText = "Test Disclaimer";




				// Populate all elements of the SS in order
				//int startRow = 1;
				//PopulateInvoiceHeader(worksheet, invoiceHeader, company, shipper, invoiceType, imagePath, ref startRow);
				//PopulateInvoiceDetailLines(worksheet, invoiceHeader, ref startRow);
				//PopulateInvoiceSummary(worksheet, invoiceHeader, invoiceType, ref startRow);
				//PopulateInvoicenote(worksheet, invoiceHeader, ref startRow);
				//PopulateInvoiceVATnote(worksheet, shipper, company, invoiceHeader, ref startRow);
				//PopulateInvoiceFootnoteData(worksheet, company, invoiceHeader, ref startRow);
				//  PopulateDisclaimer(worksheet, invoiceHeader, ref startRow);




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
		public void s610()
		{
			using(var p=OpenTemplatePackage("s610.xlsx"))
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
				//SaveWorkbook("i1317.xlsx",package);
				using(var p2=new  ExcelPackage(package.Stream)) 
				{
					var ws = p2.Workbook.Worksheets[0];
				}
			}
		}
		[TestMethod]
		public void s618()
		{
			ExcelPackage.LicenseContext = LicenseContext.Commercial;

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
					catch (Exception ex)
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
					catch (Exception ex)
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
                    //Assert.Fail("Expected no exception, but got: " + ex.Message);
                }
            }
        }
    }
}

