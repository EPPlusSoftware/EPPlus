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
using System.Diagnostics;

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

				Assert.AreEqual(10d, workbook.Worksheets["Sheet1"].Cells[1, 1].Value);

				SaveAndCleanup(p1);
			}
		}
		[TestMethod]
		public void Issue_1497_Dynamic_Array_Formulae()
		{

			//Issue: If two namedRanges (columns with Names) are calculated like "=range1 + range2" Only the first row of the ranges are calculated and the result is copied to the rest of the rows from the resultcolumn. 

#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
			var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
			using var p = OpenTemplatePackage("i1497.xlsx");

			var ws = p.Workbook.Worksheets.First();
			ws.Calculate();

			//range in range in Fomular
			Assert.AreEqual(311d, ws.Cells["C1"].Value);
			Assert.AreEqual(306d, ws.Cells["C2"].Value);

			//range1+range2 horizontal
			Assert.AreEqual(103d, ws.Cells["C3"].Value);
			Assert.AreEqual(104d, ws.Cells["C4"].Value);
			Assert.AreEqual(105d, ws.Cells["C5"].Value);
			Assert.AreEqual(106d, ws.Cells["C6"].Value);
			Assert.AreEqual(107d, ws.Cells["C7"].Value);
			Assert.AreEqual(108d, ws.Cells["C8"].Value);
			Assert.AreEqual(109d, ws.Cells["C9"].Value);
			Assert.AreEqual(110d, ws.Cells["C10"].Value);

			Assert.AreEqual(112d, ws.Cells["C12"].Value);
			Assert.AreEqual(113d, ws.Cells["C13"].Value);
			Assert.AreEqual(114d, ws.Cells["C14"].Value);

			//range3+range4 vertical
			Assert.AreEqual(101d, ws.Cells["F21"].Value);
			Assert.AreEqual(102d, ws.Cells["G21"].Value);
			Assert.AreEqual(103d, ws.Cells["H21"].Value);
			Assert.AreEqual(104d, ws.Cells["I21"].Value);
			Assert.AreEqual(105d, ws.Cells["J21"].Value);
			Assert.AreEqual(106d, ws.Cells["K21"].Value);
			Assert.AreEqual(107d, ws.Cells["L21"].Value);
			Assert.AreEqual(108d, ws.Cells["M21"].Value);
			Assert.AreEqual(109d, ws.Cells["N21"].Value);
			Assert.AreEqual(110d, ws.Cells["O21"].Value);
			Assert.AreEqual(111d, ws.Cells["P21"].Value);
			Assert.AreEqual(112d, ws.Cells["Q21"].Value);
			Assert.AreEqual(113d, ws.Cells["R21"].Value);

			//When Issue_WithRangeCalculation_IF
			Assert.AreEqual(306d, ws.Cells["H2"].Value);
			Assert.AreEqual(103d, ws.Cells["H3"].Value);
			Assert.AreEqual(104d, ws.Cells["H4"].Value);
			Assert.AreEqual(105d, ws.Cells["H5"].Value);

			Assert.AreEqual(100d, ws.Cells["I2"].Value);
			Assert.AreEqual(100d, ws.Cells["I3"].Value);
			Assert.AreEqual(100d, ws.Cells["I4"].Value);
			Assert.AreEqual(100d, ws.Cells["I5"].Value);

			Assert.AreEqual(100d, ws.Cells["J2"].Value);
			Assert.AreEqual(100d, ws.Cells["J3"].Value);
			Assert.AreEqual(100d, ws.Cells["J4"].Value);
			Assert.AreEqual(100d, ws.Cells["J5"].Value);

			Assert.AreEqual("Falsche Auswahl", ws.Cells["K2"].Value);
			Assert.AreEqual("Falsche Auswahl", ws.Cells["K3"].Value);
			Assert.AreEqual("Falsche Auswahl", ws.Cells["K4"].Value);
			Assert.AreEqual("Falsche Auswahl", ws.Cells["K5"].Value);


			//Normal
			Assert.AreEqual(198d, ws.Cells["C18"].Value);

			//String
			Assert.AreEqual("#VALUE!", ws.Cells["C19"].Value.ToString());
			Assert.AreEqual("#VALUE!", ws.Cells["C15"].Value.ToString());

			//Empty Cell
			Assert.AreEqual(100d, ws.Cells["C11"].Value);
			Assert.AreEqual(20d, ws.Cells["C20"].Value);

			//OutOfRange IF
			Assert.AreEqual("#VALUE!", ws.Cells["H1"].Value.ToString());
			Assert.AreEqual("#VALUE!", ws.Cells["I1"].Value.ToString());
			Assert.AreEqual("#VALUE!", ws.Cells["J1"].Value.ToString());
			Assert.AreEqual("Falsche Auswahl", ws.Cells["K1"].Value);
			Assert.AreEqual("#VALUE!", ws.Cells["H6"].Value.ToString());
			Assert.AreEqual("#VALUE!", ws.Cells["I6"].Value.ToString());
			Assert.AreEqual("#VALUE!", ws.Cells["J6"].Value.ToString());
			Assert.AreEqual("Falsche Auswahl", ws.Cells["K6"].Value);

			//OutOfRange Normal
			Assert.AreEqual("#VALUE!", ws.Cells["C16"].Value.ToString());
			Assert.AreEqual("#VALUE!", ws.Cells["E21"].Value.ToString());
			Assert.AreEqual("#VALUE!", ws.Cells["S21"].Value.ToString());

			//UseAGAIN
			Assert.AreEqual(206d, ws.Cells["F2"].Value);
			Assert.AreEqual(3d, ws.Cells["F3"].Value);
			Assert.AreEqual(4d, ws.Cells["F4"].Value);
			Assert.AreEqual(5d, ws.Cells["F5"].Value);
			//UseIFAGAIN
			Assert.AreEqual(306d, ws.Cells["M2"].Value);
			Assert.AreEqual(103d, ws.Cells["M3"].Value);
			Assert.AreEqual(104d, ws.Cells["M4"].Value);
			Assert.AreEqual(105d, ws.Cells["M5"].Value);
			Assert.AreEqual("#VALUE!", ws.Cells["M6"].Value.ToString());


			//Check if something in if is fixed wrong
			Assert.AreEqual(2d, ws.Cells["F11"].Value);
			Assert.AreEqual(1d, ws.Cells["F12"].Value);
		}
		[TestMethod]
		public void s701()
		{
			using (var package = OpenTemplatePackage("s701.xlsx"))
			{
				var wk = package.Workbook.Worksheets[0];
				Debug.WriteLine($"Open Cell B5 Value:{wk.Cells["B5"].Value}");

        Debug.WriteLine($"Open Cell A5 Formula:{wk.Cells["A5"].Formula}");
        Debug.WriteLine($"Open Cell A5 Value:{wk.Cells["A5"].Value}");

				package.Workbook.Calculate();

				wk.InsertRow(2, 4);
				wk.Cells["B5"].Value = "Error B5";

				Debug.WriteLine($"Before recalculate Cell B9 Value:{wk.Cells["B9"].Value}");

				Debug.WriteLine($"Before recalculate Cell A9 Formula:{wk.Cells["A9"].Formula}");

				Debug.WriteLine($"Before recalculate Cell A9 Value:{wk.Cells["A9"].Value}");

				package.Workbook.Calculate(x => x.CacheExpressions = false); // get value to original row before insert row

				Debug.WriteLine($"After Cell B9 Value:{wk.Cells["B9"].Value}");

				Debug.WriteLine($"After Cell A9 Formula:{wk.Cells["A9"].Formula}");

				Debug.WriteLine($"After Cell A9 Value:{wk.Cells["A9"].Value}");
			}
		}
		[TestMethod]
		public void i1540()
		{
			using (var p = OpenPackage("i1540.xlsx",true))
			{
				var ws = p.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1"].Value = "A";
                ws.Cells["A2"].Value = "B";
                ws.Cells["A3"].Value = "C";
                ws.Cells["B1:B3"].FillNumber(1, 1);
                ws.Cells["C1:C3"].FillNumber(10, 10);
				ws.Cells["E1"].Formula = "SUM(If(A:A=\"A\",B:B,C:C))";							//Should be set as an array formula
                ws.Cells["E2"].Formula = "SUM(If(A1:A3=\"A\",B1:B3,C1:C3))";                    //Should be set as an array formula
                ws.Cells["F1"].Formula = "SUM(If(A:A=\"A\",B:B,C:C))";							//Should be set as an array formula
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
        [TestMethod]
        public void i1566()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                /* 
                This reference to a custom function is a simulation of my use-case.
                It doesn't appear to matter what the formula is, it just has to be set to something
                ws.Cells["A3"].Formula = "1"; // this works just as well as "@SomeCustomVbaFunction(A1,A2)"
                */
                ws.Cells["A3"].Formula = "@SomeCustomVbaFunction(A1,A2)";
                /* 
                 * clear the formulas so that EPPlus doesn't go looking for SomeCustomVbaFunction
                 I have purposefully chosen not to implement this function as a class extending ExcelFunction                
                */
                ws.Cells["A3"].ClearFormulas();
                //ws.Cells["A3"].Formula = "0"; //This may be a workaround for now
                ws.Cells["A3"].Value = "2000";
                ws.Cells["A4"].Formula = "ROUNDUP(A3/1609.334,0)";

                ws.Calculate();
                Assert.AreEqual(2D, ws.Cells["A4"].Value);

            }
        }

		[TestMethod]
		public void i1671()
		{
			using var package = new ExcelPackage();
			var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
            var sheet2 = package.Workbook.Worksheets.Add("Sheet2");

			sheet1.Cells["A1"].Value = "h1";
			sheet1.Cells["B1"].Value = "h2";
			sheet1.Cells["C1"].Value = "h3";
			sheet1.Cells["A2"].Value = "a1";
            sheet1.Cells["B2"].Formula = "VLOOKUP($A2,Sheet2!$A:$B,2,FALSE)";
            sheet1.Cells["C2"].Formula = "VLOOKUP($A2,Sheet2!$A:$C,3,FALSE)";

            sheet2.Cells["A1"].Value = "a1";
            sheet2.Cells["B1"].Value = "b1";
            sheet2.Cells["C1"].Value = "c1";
            sheet2.Cells["A2"].Value = "a2";
			sheet2.Cells["B2"].Value = "b2";
			sheet2.Cells["C2"].Value = "c2";

			Assert.IsNull(sheet1.Cells["B2"].Value);

			sheet1.Calculate();

			Assert.AreEqual("b1", sheet1.Cells["B2"].Value);
			Assert.AreEqual("c1", sheet1.Cells["C2"].Value);		
        }
        [TestMethod]
        public void Issue1696()
        {
            using (var wb = OpenTemplatePackage("i1696-1.xlsx"))
            {
                wb.Workbook.Worksheets.Copy("template", "Test-Copy");
                wb.Workbook.Calculate();
                wb.Workbook.Worksheets.Delete("template");

                wb.Workbook.Calculate();
            }

            using (var wb = OpenTemplatePackage("i1696-2.xlsx"))
            {
				wb.Compatibility.IsWorksheets1Based = true;
				wb.Workbook.Worksheets.Copy("template", "Test-Copy");
                wb.Workbook.Calculate();
                wb.Workbook.Worksheets.Delete("template");

                wb.Workbook.Calculate();
            }
        }
		[TestMethod]
        public void i1708()
        {
            using (var package = OpenPackage("i1708.xlsx"))
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                package.Compatibility.IsWorksheets1Based = true;

                sheet1.Cells["C3"].Formula = @"IFERROR(IF(OR(H3="""",I3="""",E3=0),""N/A"",IF(J3<>"""",INDEX($G$1:$J$1,MATCH(TRUE,INDEX(ABS(G3:J3-E3)=MIN(INDEX(ABS(G3:J3-E3),,)),,),0)),INDEX($G$1:$I$1,MATCH(TRUE,INDEX(ABS(G3:I3-E3)=MIN(INDEX(ABS(G3:I3-E3),,)),,),0)))),"""")";
                sheet1.Cells["E3"].Value = 25;

                sheet1.Cells["G1"].Value = "one";
                sheet1.Cells["H1"].Value = "two";
                sheet1.Cells["I1"].Value = "three";
                sheet1.Cells["J1"].Value = "four";

                sheet1.Cells["G3"].Value = 10;
                sheet1.Cells["H3"].Value = 20;
                sheet1.Cells["I3"].Value = 30;
                sheet1.Cells["J3"].Value = 40;

                package.Workbook.Calculate();
                Assert.AreEqual("two", sheet1.Cells["C3"].Value);
            }
        }

		[TestMethod]
		public void i1729()
		{
			using var package = new ExcelPackage();
			var worksheet = package.Workbook.Worksheets.Add("Sheet1");
			worksheet.Cells["A1"].Value = "A";
			worksheet.Cells["A2"].Formula = "VLOOKUP(1,B1:C2,2,FALSE)"; //Return #N/A
            worksheet.Cells["A3"].Value = "B";
            worksheet.Cells["A4"].Formula = "TEXTJOIN(\"\",TRUE,A1:A3)";
            worksheet.Cells["A5"].Formula = "TEXTJOIN(\"\",TRUE,A1,A2,A3)";
            worksheet.Cells["A6"].Formula = "CONCAT(A1:A3)";
            worksheet.Cells["A7"].Formula = "CONCAT(A1,A2,A3)";
			worksheet.Calculate();
			var a4 = worksheet.Cells["A4"].Value;
            var a5 = worksheet.Cells["A5"].Value;
            var a6 = worksheet.Cells["A6"].Value;
            var a7 = worksheet.Cells["A7"].Value;

			var naError = ExcelErrorValue.Create(eErrorType.NA);

			Assert.AreEqual(naError, a4);
			Assert.AreEqual(naError, a5);
			Assert.AreEqual(naError, a6);
			Assert.AreEqual(naError, a7);
        }
    }
}