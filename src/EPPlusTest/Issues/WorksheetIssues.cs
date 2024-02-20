using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
namespace EPPlusTest
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




				string longText = ""; // Our long string
				int maxLineLength = 140; // Maximum length of each line, adjust as needed
										 //var lines = SplitStringIntoLines(longText, maxLineLength);


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
	}
}
