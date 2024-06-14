using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing;
using System.IO;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.Issues
{
	[TestClass]
	public class FormulaCalculationIssues : TestBase
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
		public void I1228()
		{
			using (var pck = new ExcelPackage())
			{
				using (var pckTemplate = OpenTemplatePackage("MyIssue.xlsx"))
				{
					pck.Workbook.Worksheets.Add("Foo", pckTemplate.Workbook.Worksheets[1]);
				}

				pck.Workbook.Calculate(x => x.AllowCircularReferences = true);
			}
		}
		[TestMethod]
		public void I1229()
		{
			using (var p = OpenPackage("XLOOKUP.xlsx"))
			{
				var ws = p.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1:A5"].Formula = "XLOOKUP(B1,$C$1:$C$5,$D$1:$D$5,0)";
				ws.Cells["E1"].Formula = "XLOOKUP(B1:B5,$C$1:$C$5,$D$1:$D$5,0)";

				ws.Cells["B1"].Value = 1;
				ws.Cells["B2"].Value = 2;
				ws.Cells["B3"].Value = 3;
				ws.Cells["B4"].Value = 5;
				ws.Cells["B5"].Value = 4;

				ws.Cells["C1"].Value = 1;
				ws.Cells["C2"].Value = 2;
				ws.Cells["C3"].Value = 3;
				ws.Cells["C4"].Value = 5;
				ws.Cells["C5"].Value = 4;

				ws.Cells["D1"].Value = 10;
				ws.Cells["D2"].Value = 12;
				ws.Cells["D3"].Value = 13;
				ws.Cells["D4"].Value = 14;
				ws.Cells["D5"].Value = 15;


				p.Workbook.Calculate();

				Assert.AreEqual(10, ws.Cells["A1"].Value);
				Assert.AreEqual(12, ws.Cells["A2"].Value);
				Assert.AreEqual(13, ws.Cells["A3"].Value);
				Assert.AreEqual(14, ws.Cells["A4"].Value);
				Assert.AreEqual(15, ws.Cells["A5"].Value);

				Assert.AreEqual(10, ws.Cells["E1"].Value);
				Assert.AreEqual(12, ws.Cells["E2"].Value);
				Assert.AreEqual(13, ws.Cells["E3"].Value);
				Assert.AreEqual(14, ws.Cells["E4"].Value);
				Assert.AreEqual(15, ws.Cells["E5"].Value);

			}
		}
		[TestMethod]
		public void ImplicitIntersection_ColumnReference()
		{
			using (var pck = new ExcelPackage())
			{
				var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["E2"].Value = 12;
				sheet1.Cells["E3"].Value = 23;
				sheet1.Cells["E4"].Value = 34;
				sheet1.Cells["E5"].Value = 45;

				sheet1.Cells["C3"].Formula = "E:E";
				sheet1.Cells["C4"].Formula = "E1:E5";

				sheet1.Cells["C3:C4"].UseImplicitItersection = true;

				pck.Workbook.Calculate();

				Assert.AreEqual(23D, sheet1.Cells["C3"].GetValue<double>());
				Assert.AreEqual(34D, sheet1.Cells["C4"].GetValue<double>());
			}
		}
		[TestMethod]
		public void i1234()
		{
			using (var p = OpenTemplatePackage("i1234.xlsx"))
			{
				SaveAndCleanup(p);
			}
		}		

		[TestMethod]
		public void SubtractWorksheetReference()
		{
			const string MinusQuoteFormula = "10-'Sheet A'!A1";
			const string SheetA = "Sheet A";

			using var setupPackage = new ExcelPackage();
			setupPackage.Workbook.Worksheets.Add(SheetA);
			var sheetA = setupPackage.Workbook.Worksheets[SheetA];
			sheetA.Cells[1, 1].Value = 3;
			sheetA.Cells[1, 2].Formula = MinusQuoteFormula;

			var stream = new MemoryStream();
			setupPackage.SaveAs(stream);

			using var testPackage = new ExcelPackage(stream);
			string savedMinusQuoteFormula = testPackage.Workbook.Worksheets[SheetA].Cells[1, 2].Formula;
			Assert.AreEqual(MinusQuoteFormula, savedMinusQuoteFormula);
		}

		[TestMethod]
		public void s568()
		{
			using (var p = OpenTemplatePackage("s568.xlsx"))
			{
				p.Workbook.Calculate();
				SaveAndCleanup(p);
			}
		}
		[TestMethod]
		public void i1244()
		{
			using (var p = OpenTemplatePackage("i1245.xlsx"))
			{
				var wbk = p.Workbook;
				var sht = wbk.Worksheets["TestSheet"];

				// Call calculate
				wbk.Calculate();

				// Check everything is initially in order
				Assert.AreEqual(1.0, sht.Cells["B3"].Value);
				Assert.AreEqual(2.0, sht.Cells["C3"].Value);
				Assert.AreEqual(2.0, sht.Cells["B4"].Value);
				Assert.AreEqual(4.0, sht.Cells["C4"].Value);

				// Update the value of two cells
				sht.Cells["B3"].Value = 500.0;
				sht.Cells["B4"].Value = 500.0;


				var form1 = sht.Cells["C3"].Formula;
				var form2 = sht.Cells["C4"].Formula;

				wbk.Calculate();

				Assert.AreEqual(1000.0, sht.Cells["C3"].Value);
				Assert.AreEqual(1000.0, sht.Cells["C4"].Value);

				SaveAndCleanup(p);
			}
		}
		[TestMethod]
		public void i1335()
		{
			var formula = "SUBTOTAL(109, Name1 Name2)";
			var tokens = SourceCodeTokenizer.Default_KeepWhiteSpaces.Tokenize(formula);

			Assert.AreEqual(9, tokens.Count);
			Assert.AreEqual(TokenType.WhiteSpace, tokens[4].TokenType);
			Assert.AreEqual(TokenType.Operator, tokens[6].TokenType);
			Assert.AreEqual("isc", tokens[6].Value);
		}
		[TestMethod]
		public void s637()
		{
			using (var p = OpenTemplatePackage("s637.xlsx"))
			{
				SaveAndCleanup(p);
			}
		}
        [TestMethod]
        public void CalcError()
		{
			using (var package = OpenTemplatePackage("calc.xlsx"))
			{
				var summary =
				package.Workbook.Worksheets["Summary"];
				ExcelCalculationOption eco = new();
				eco.AllowCircularReferences = true;
				eco.CacheExpressions = false;
				var original = summary.Cells["M22"].Value;
				package.Workbook.Calculate(eco);
				Assert.AreEqual(42354.210446, (double)summary.Cells["M22"].Value, 0.000001);
			}
        }
        [TestMethod]
        public void s681()
		{
			using (var p1 = OpenTemplatePackage("s681-bad.xlsx"))
			{
				ExcelWorkbook workbook = p1.Workbook;
				SaveAndCleanup(p1);
				//SaveWorkbook("s681Good.xlsx",p1);
            }

    //        using (var p2 = OpenPackage("s681Good.xlsx"))
    //        {
    //            ExcelWorkbook workbook = p2.Workbook;
				//SaveWorkbook("s681Bad.xlsx", p2);

    //        }
        }
		[TestMethod]
		public void s684()
		{
            using (var p1 = OpenTemplatePackage("s684.xlsx"))
            {
				p1.Compatibility.IsWorksheets1Based = true;
                ExcelWorkbook workbook = p1.Workbook;
                workbook.Calculate();
				Assert.AreEqual(7d, workbook.Worksheets["Sheet1"].Cells[1, 1].Value);
                
				workbook.Worksheets.First().Cells[2, 1].Value = 4;
                workbook.Calculate();
                
				Assert.AreEqual(10d, workbook.Worksheets["Sheet1"].Cells[1,1].Value);

                SaveAndCleanup(p1);
            }
        }
    }
}