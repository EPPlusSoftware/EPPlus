using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Sorting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

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
            using (var p = OpenPackage("XLOOKUP.xlsx", true))
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

                SaveWorkbook("XLOOKUP.xlsx", p);

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
        public void FormulaDemo()
        {
            using (var p1 = OpenTemplatePackage("s684.xlsx"))
            {
                ExcelWorkbook workbook = p1.Workbook;
                workbook.Worksheets[0].Cells["A1"].Calculate();
            }
        }

        [TestMethod]
        public void s684()
        {
            using (var p1 = OpenTemplatePackage("s684.xlsx"))
            {
                p1.Compatibility.IsWorksheets1Based = true;
                ExcelWorkbook workbook = p1.Workbook;
                workbook.Calculate();
                Assert.AreEqual(8.333333d, (double)workbook.Worksheets["Sheet1"].Cells[1, 1].Value, 0.00001);

                workbook.Worksheets.First().Cells[2, 1].Value = 4;
                workbook.Calculate();

                Assert.AreEqual(11.333333d, (double)workbook.Worksheets["Sheet1"].Cells[1, 1].Value, 0.00001);

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
        [TestMethod]
        public void i1748()
        {
            using var package = new ExcelPackage();
            var formula = "SUMIF($I$3:$L$3,1,INDEX($I:$I,ROW()):INDEX($L:$L,ROW()))";
            var formulaLong = "IF(COLUMN()-COLUMN($J23)>(COUNTA('#CompaniesAndConsolidations'!$A:$A)+1),0,IF(K$5=\"TopConsolidation\",SUMIFS(INDEX(23:23,1,COLUMN()+1):INDEX(23:23,1,COLUMN($L23)),INDEX($7:$7,1,COLUMN()+1):INDEX($7:$7,1,COLUMN($L23)),K$2,INDEX($6:$6,1,COLUMN()+1):INDEX($6:$6,1,COLUMN($L23)),FALSE),IF(K$5=\"SubConsolidation\",SUMIFS(INDEX(23:23,1,COLUMN()+1):INDEX(23:23,1,COLUMN($L23)),INDEX($8:$8,1,COLUMN()+1):INDEX($8:$8,1,COLUMN($L23)),K$2,INDEX($6:$6,1,COLUMN()+1):INDEX($6:$6,1,COLUMN($L23)),FALSE),IF(K$5=\"DivisionalConsolidation\",SUMIFS(INDEX(23:23,1,COLUMN()+1):INDEX(23:23,1,COLUMN($L23)),INDEX($9:$9,1,COLUMN()+1):INDEX($9:$9,1,COLUMN($L23)),K$2,INDEX($6:$6,1,COLUMN()+1):INDEX($6:$6,1,COLUMN($L23)),FALSE),-SUMIFS('#TrialBalance_CY'!$E:$E,'#TrialBalance_CY'!$A:$A,K$2,'#TrialBalance_CY'!$G:$G,\"IncomeStatement\")))))";
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Formula = formula;
            ws.Cells["A2"].Formula = formulaLong;
            ws.InsertRow(1, 1);

            var formulaInserted = "SUMIF($I$4:$L$4,1,INDEX($I:$I,ROW()):INDEX($L:$L,ROW()))";

            Assert.AreEqual(formulaInserted, ws.Cells["A2"].Formula);

            var formulaLongInserted = "IF(COLUMN()-COLUMN($J24)>(COUNTA('#CompaniesAndConsolidations'!$A:$A)+1),0,IF(K$6=\"TopConsolidation\",SUMIFS(INDEX(24:24,1,COLUMN()+1):INDEX(24:24,1,COLUMN($L24)),INDEX($8:$8,1,COLUMN()+1):INDEX($8:$8,1,COLUMN($L24)),K$3,INDEX($7:$7,1,COLUMN()+1):INDEX($7:$7,1,COLUMN($L24)),FALSE),IF(K$6=\"SubConsolidation\",SUMIFS(INDEX(24:24,1,COLUMN()+1):INDEX(24:24,1,COLUMN($L24)),INDEX($9:$9,1,COLUMN()+1):INDEX($9:$9,1,COLUMN($L24)),K$3,INDEX($7:$7,1,COLUMN()+1):INDEX($7:$7,1,COLUMN($L24)),FALSE),IF(K$6=\"DivisionalConsolidation\",SUMIFS(INDEX(24:24,1,COLUMN()+1):INDEX(24:24,1,COLUMN($L24)),INDEX($10:$10,1,COLUMN()+1):INDEX($10:$10,1,COLUMN($L24)),K$3,INDEX($7:$7,1,COLUMN()+1):INDEX($7:$7,1,COLUMN($L24)),FALSE),-SUMIFS('#TrialBalance_CY'!$E:$E,'#TrialBalance_CY'!$A:$A,K$3,'#TrialBalance_CY'!$G:$G,\"IncomeStatement\")))))";

            Assert.AreEqual(formulaLongInserted, ws.Cells["A3"].Formula);
        }
        [TestMethod]
        public void s780()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Formula = "IF(B10+C10=0,\"\",B10+C10)";
            ws.Cells["A2"].Formula = "IF(B10+C10=0,\" \",B10+C10)";
            ws.Calculate();

            Assert.AreEqual("", ws.Cells["A1"].Value);
            Assert.AreEqual(" ", ws.Cells["A2"].Value);
        }
        [TestMethod]
        public void i1766()
        {
            var ep = OpenTemplatePackage("i1766.xlsx");

            Assert.AreEqual(ep.Workbook.CalcMode, ExcelCalcMode.Automatic);

            var wr = ep.Workbook.Names["Width"];
            var hr = ep.Workbook.Names["Height"];
            var ar = ep.Workbook.Names["Area"];

            wr.Worksheet.SetValue(wr.Address, 5);
            hr.Worksheet.SetValue(hr.Address, 7);

            ar.Calculate(); //no matter if we do ar.Calculate() or not, the value in the range will not update and remains 200.
                            //ar.Worksheet.Calculate(); manual calculate on worksheet will solve the problem

            var area = ar.Value;
            var area2 = ar.Worksheet.Cells[ar.Address].Value;

            Assert.AreEqual(area, 35D);
            Assert.AreEqual(area, area2);
        }

        [TestMethod]
        public void RedYellowGreenShouldNotCreateCorruptWorkbookReproduce()
        {
            using (var p = OpenPackage("RedYellowGreenUncorrupt.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("sheet1");

                sheet.Cells["D27"].Value = -1;
                sheet.Cells["D28"].Value = 0;
                sheet.Cells["D29"].Value = 1;

                sheet.Cells["E27"].Value = "RedL";
                sheet.Cells["E28"].Value = "YellowL";
                sheet.Cells["E29"].Value = "GreenL";


                sheet.Cells["F27"].Value = "RED";
                sheet.Cells["F28"].Value = "Yellow";
                sheet.Cells["F29"].Value = "Green";

                sheet.Cells["F27:F29"].Formula = "IFS(E27=\"RedL\",\"RED\",E27=\"YellowL\",\"Yellow\",E27=\"GreenL\",\"Green\")";


                var firstRange = sheet.Cells["D27:F30"];

                var options = RangeSortOptions.Create();
                options.SortLeftToRightBy.Row(0);
                firstRange.Sort(options);

                sheet.Calculate();

                SaveAndCleanup(p);
            }
            using (var p = OpenPackage("RedYellowGreenUncorrupt.xlsx"))
            {
                var sheet = p.Workbook.Worksheets.First();

                var firstRange = sheet.Cells["D27:F30"];

                var options = RangeSortOptions.Create();
                options.SortLeftToRightBy.Row(0);
                firstRange.Sort(options);

                sheet.Calculate();

                var outFile = GetOutputFile("", "RedYellowGreenCorrupt.xlsx");
                p.SaveAs(outFile.FullName);
            }
        }

        [TestMethod]
        public void RedYellowGreenShouldNotCreateCorruptWorkbook()
        {
            using (var p = OpenTemplatePackage("RedYellowGreen.xlsx"))
            {
                var sheet = p.Workbook.Worksheets.First();

                var firstRange = sheet.Cells["D27:F30"];

                var options = RangeSortOptions.Create();
                options.SortLeftToRightBy.Row(0);
                firstRange.Sort(options);

                sheet.Calculate();

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void RedYellowGreen_NoPrevDimension()
        {
            using (var p = OpenPackage("RedYellowGreen_NoDim.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("NewWs");

                sheet.Cells["D27"].Value = 1;
                sheet.Cells["D28"].Value = 0;
                sheet.Cells["D29"].Value = -1;

                sheet.Cells["E27"].Value = "GreenLight";
                sheet.Cells["E28"].Value = "YellowLight";
                sheet.Cells["E29"].Value = "RedLight";

                sheet.Cells["F27:F29"].Formula = "IFS(E27=\"RedLight\",\"Red\",E27=\"YellowLight\",\"Yellow\",E27=\"GreenLight\",\"Green\")";

                var firstRange = sheet.Cells["D20:F30"];

                firstRange.Sort(column: 0);

                Assert.AreEqual(-1, sheet.Cells["D27"].Value, "D27 wasn't -1 as expected");
                Assert.AreEqual(0, sheet.Cells["D28"].Value, "D28 wasn't 0 as expected");
                Assert.AreEqual(1, sheet.Cells["D29"].Value);

                Assert.AreEqual("RedLight", sheet.Cells["E27"].Value);
                Assert.AreEqual("YellowLight", sheet.Cells["E28"].Value);
                Assert.AreEqual("GreenLight", sheet.Cells["E29"].Value);

                sheet.Calculate();

                Assert.AreEqual("Red", sheet.Cells["F27"].Value);
                Assert.AreEqual("Yellow", sheet.Cells["F28"].Value);
                Assert.AreEqual("Green", sheet.Cells["F29"].Value);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void RedYellowGreen_PrevDimension()
        {
            using (var p = OpenPackage("RedYellowGreen_Dim.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("NewWs");

                // add a random value above the sorted range
                sheet.Cells["D14"].Value = 3;

                sheet.Cells["D27"].Value = 1;
                sheet.Cells["D28"].Value = 0;
                sheet.Cells["D29"].Value = -1;

                sheet.Cells["E27"].Value = "GreenLight";
                sheet.Cells["E28"].Value = "YellowLight";
                sheet.Cells["E29"].Value = "RedLight";

                sheet.Cells["F27:F29"].Formula = "IFS(E27=\"RedLight\",\"Red\",E27=\"YellowLight\",\"Yellow\",E27=\"GreenLight\",\"Green\")";

                var firstRange = sheet.Cells["D20:F30"];

                firstRange.Sort(column: 0);

                Assert.AreEqual(-1, sheet.Cells["D20"].Value, "D27 wasn't -1 as expected");
                Assert.AreEqual(0, sheet.Cells["D21"].Value, "D28 wasn't 0 as expected");
                Assert.AreEqual(1, sheet.Cells["D22"].Value);

                Assert.AreEqual("RedLight", sheet.Cells["E20"].Value);
                Assert.AreEqual("YellowLight", sheet.Cells["E21"].Value);
                Assert.AreEqual("GreenLight", sheet.Cells["E22"].Value);

                sheet.Calculate();

                Assert.AreNotEqual("", sheet.Cells["F20"].Formula, "Formula in F20 was empty");
                Assert.AreEqual("Red", sheet.Cells["F20"].Value);
                Assert.AreEqual("Yellow", sheet.Cells["F21"].Value);
                Assert.AreEqual("Green", sheet.Cells["F22"].Value);

                Assert.IsNull(sheet.Cells["D27"].Value);
                Assert.IsNull(sheet.Cells["D28"].Value);
                Assert.IsNull(sheet.Cells["D29"].Value);
                Assert.IsNull(sheet.Cells["E27"].Value);
                Assert.IsNull(sheet.Cells["E28"].Value);
                Assert.IsNull(sheet.Cells["E29"].Value);
                Assert.AreEqual("", sheet.Cells["F27"].Formula, "Formula still set in F27");
                Assert.AreEqual("", sheet.Cells["F28"].Formula, "Formula still set in F28");
                Assert.AreEqual("", sheet.Cells["F29"].Formula, "Formula still set in F29");

                SaveAndCleanup(p);
            }
        }


        [TestMethod]
        public void S809()
        {
            using (var p = OpenTemplatePackage("s809.xlsx"))
            {
                var sheet = p.Workbook.Worksheets.First();

                sheet.Cells.Sort(column: 0);
                sheet.Calculate();

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void Sc809()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 3;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 1;
            sheet.Cells["B1"].Value = "All Motor";
            sheet.Cells["B2"].Value = "All Rail";
            sheet.Cells["B3"].Value = "All Rail";

            sheet.Cells["C1:C3"].Formula = "IFS(B1=\"All Rail\",\"Rail\",B1=\"All Motor\",\"Road\",B1=\"All Barge\",\"Barge\")";

            sheet.Cells.Sort(column: 0);
            sheet.Calculate();
            Assert.AreEqual(1, sheet.Cells["A1"].Value);
            Assert.AreEqual(2, sheet.Cells["A2"].Value);
            Assert.AreEqual(3, sheet.Cells["A3"].Value);
            Assert.AreEqual("Rail", sheet.Cells["C1"].Value);
            Assert.AreEqual("Rail", sheet.Cells["C2"].Value);
            Assert.AreEqual("Road", sheet.Cells["C3"].Value);
            SaveWorkbook("Sc809_Output_NotSorted.xlsx", p);
        }

        [TestMethod]
        public void Issue1828()
        {
            using var p = OpenTemplatePackage("Issue1828.xlsx");
            var sheet = p.Workbook.Worksheets.First();

            sheet.Cells.Sort(column: 0);
            sheet.Calculate();

            SaveAndCleanup(p);
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
        public void Issue1687()
        {
            using var p = OpenTemplatePackage("i1687.xlsx");
            var sheet = p.Workbook.Worksheets.First();
            sheet.Cells["D5"].ClearFormulaValues();
            sheet.Calculate();
            Assert.AreEqual(44D, sheet.Cells["D5"].Value);
            SaveAndCleanup(p);

        }
        public void i1930()
        {
            using var p = OpenPackage("i1930.xlsx");
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            LoadTestdata(ws);

            ws.Cells["G1"].Formula = "LET(var1,Table1[[#This Row],[Col1]],var2, Table1[[#This Row],[Col2]],var1 + var2);";

        }
        [TestMethod]
        public void LetTwice()
        {
            using (var pck = OpenPackage("LetFunction_Twice.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("LET params");
                var table = sheet.Tables.Add(sheet.Cells["D1:E10"], "Table1");
                sheet.Cells["D2:D10"].FillNumber(1, 1);
                sheet.Cells["E2:E10"].FillNumber(2, 2);
                table.SyncColumnNames(OfficeOpenXml.Table.ApplyDataFrom.ColumnNamesToCells);
                sheet.Cells["A2"].Formula = "LET(var1, Table1[[#This Row],[Column1]], var2, Table1[[#This Row],[Column2]], var1 + var2)";
                sheet.Cells["A3"].Formula = "LET(var1, Table1[[#This Row],[Column1]], var2, Table1[[#This Row],[Column2]], var1 + var2)";
                sheet.Calculate();
                Assert.AreEqual(3D, sheet.Cells["A2"].Value);
                Assert.AreEqual(6D, sheet.Cells["A3"].Value);
                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void s842()
        {
            using var p = OpenTemplatePackage("sapreport broken.xlsx");
            p.Workbook.CalcMode = ExcelCalcMode.Automatic;
            var opt = new OfficeOpenXml.FormulaParsing.ExcelCalculationOption
            {
                PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel
            };
            p.Workbook.Calculate();
            var ws = p.Workbook.Worksheets.First();
            var val = ws.Cells["A1"].Value;
        }
        [TestMethod]
        public void s846()
        {
            using var p = OpenTemplatePackage("s846.xlsx");
            p.Workbook.Calculate();
            var ws = p.Workbook.Worksheets["Calculation sheet"];
            ws.Calculate();
            Assert.AreEqual(321732.45, ws.Cells["H11"].Value);
            var ws2 = p.Workbook.Worksheets["aico data"];
            ws2.Calculate();
            Assert.AreEqual(93515.9075, ws2.Cells["D32"].Value);
        }
        [TestMethod]
        public void s851()
        {
            using var excelPackage = OpenTemplatePackage("s851.xlsx");
            var sheet = excelPackage.Workbook.Worksheets.First();

            // Act
            sheet.Cells.Sort(column: 0); // NullReferenceException

            // Assert 1
            Assert.AreEqual("VLOOKUP(B2,B1:C1,1,FALSE)", sheet.Cells["C2"].Formula);
        }
        [TestMethod]
        public void s851_desc()
        {
            using var excelPackage = OpenTemplatePackage("s851.xlsx");
            var sheet = excelPackage.Workbook.Worksheets.First();

            // Act
            sheet.Cells.Sort(column: 0, true); // NullReferenceException

            // Assert 1
            Assert.AreEqual("VLOOKUP(B1,#REF!,1,FALSE)", sheet.Cells["C1"].Formula);
        }
        [TestMethod]
        public void s853()
        {
            using var excelPackage = OpenTemplatePackage("s853.xlsx");
            var sheet = excelPackage.Workbook.Worksheets["Aico data"];

            // Act
            sheet.Cells["AH61"].Calculate();
            Assert.AreEqual("AH61:AH221", sheet.Cells["AH61"].FormulaRange.Address);
            Assert.AreEqual("DBDBEUR2", sheet.Cells["AH61"].Value);
            Assert.AreEqual("255007727", sheet.Cells["AH70"].Value);

            sheet.Cells["AI61"].Calculate();

            Assert.AreEqual("AI61:AI221", sheet.Cells["AI61"].FormulaRange.Address);
            Assert.AreEqual(12273.13, sheet.Cells["AI61"].Value);
            Assert.AreEqual(-472.69, sheet.Cells["AI70"].Value);
        }
        [TestMethod]
        public void s858()
        {
            using var p1 = OpenTemplatePackage("s858-1.xlsx");
            var ws1 = p1.Workbook.Worksheets["Aico Data"];
            ws1.Calculate();
            var result1 = ws1.Cells["D55"].Value;
            Assert.AreEqual(265509.38, result1);

            using var p2 = OpenTemplatePackage("s858-2.xlsx");
            var ws2 = p2.Workbook.Worksheets["Aico data"];
            ws2.Calculate();
            var result2 = ws2.Cells["E55"].Value;

            Assert.AreEqual(12977661.57, result2);
        }
        [TestMethod]
        public void i2012_Alt_Minimized_Col()
        {
            using (var pck = OpenTemplatePackage("SortColumn.xlsx"))
            {
                var ws = pck.Workbook.Worksheets[0];

                ws.Cells["A13"].Formula = "SORT(A1:BB4,2,1,TRUE)";

                ws.Cells["A13"].Calculate();
                ws.ClearFormulas();

                List<string> cellValues = new();
                foreach (var cell in ws.Cells["A7:BB10"])
                {
                    cellValues.Add(cell.Text);
                }

                int i = 0;
                foreach (var cell in ws.Cells["A13:BB16"])
                {
                    Assert.AreEqual(cell.Text, cellValues[i]);
                    i++;
                }

                SaveAndCleanup(pck);
            }
        }
        [TestMethod]
        public void i2012_Alt_Minimized()
        {
            using (var pck = OpenTemplatePackage("MinimizedSort.xlsx"))
            {
                var ws = pck.Workbook.Worksheets.Add("EpplusSort");

                ws.Cells["A1"].Formula = "SORT(Data!A1:D3,2,1,FALSE)";
                ws.Cells["G1"].Formula = "SORT(Data!A5:D12,2,1,FALSE)";

                ws.Calculate();
                ws.ClearFormulas();

                ws.Cells["G1:J8"].CopyTranspose(ws.Cells["L1"]);

                ws.Cells["U1"].Formula = "SORT(L1:S4,3,1,TRUE)";
                //ws.Cells["U1"].IsArrayFormula = false;

                ws.Calculate();
                ws.ClearFormulas();

                ws.Cells["AH1"].Formula = "SORT(L1:S4,3,1,TRUE)";

                ws.Calculate();

                //ws.Cells["U7"].ClearFormulas();

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void i2012_Alt()
        {
            using (var pck = OpenTemplatePackage("BrokeOutProblem.xlsx"))
            {
                var wsExcel = pck.Workbook.Worksheets[0];
                var wsEpplus = pck.Workbook.Worksheets.Add("EpplusCalc");

                wsEpplus.Cells["A1"].Formula = wsExcel.Cells["A1"].Formula;

                int i = 0;
                wsEpplus.Cells["H1"].Formula = "SORT(Sheet2!A1:D54,{2,4},1,FALSE)";
                wsEpplus.Calculate();

                wsEpplus.ClearFormulas();

                foreach (var cell in wsEpplus.Cells["B1:B55"])
                {
                    if (cell.Text != wsExcel.Cells[cell.Start.Row, cell.Start.Column].Text)
                    {
                        cell.EntireRow.Style.Fill.SetBackground(System.Drawing.Color.DarkRed);
                    }
                    i++;
                }

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void i2012()
        {
            using (var pck = OpenTemplatePackage("i2012.xlsx"))
            {
                var ws = pck.Workbook.Worksheets["Summary by Contract"];
                var str = ws.Cells["A12"].Formula;

                List<string> cellValues = new();
                foreach (var cell in ws.Cells["D12:D100"])
                {
                    cellValues.Add(cell.Text);
                }

                ws.Cells["A12:D200"].Clear();

                ws.Cells["A12"].Formula = "IF('Contract Details'!F12 = \"[none]\", {\"[none]\",\"\",\"\",\"\"},\r\n _xlfn.VSTACK(\r\n_xlfn._xlws.SORT(_xlfn.VSTACK(\r\n_xlfn.HSTACK(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!A12),_xlfn.ANCHORARRAY('Contract Details'!B12))), REPT(name_Account_GL_Current,_xlfn.SEQUENCE(COUNTA(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!B12)))),1,1,0)), REPT(\"Current\",_xlfn.SEQUENCE(COUNTA(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!B12)))),1,1,0))),\r\n_xlfn.HSTACK(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!A12),_xlfn.ANCHORARRAY('Contract Details'!B12))), REPT(name_Account_GL_NonCurrent,_xlfn.SEQUENCE(COUNTA(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!B12)))),1,1,0)), REPT(\"Non-Current\",_xlfn.SEQUENCE(COUNTA(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!B12)))),1,1,0))),\r\n_xlfn.HSTACK(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!A12),_xlfn.ANCHORARRAY('Contract Details'!B12))), REPT(name_Account_GL_Total,_xlfn.SEQUENCE(COUNTA(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!B12)))),1,1,0)), REPT(\"🔼 Sub-Total\",_xlfn.SEQUENCE(COUNTA(_xlfn.UNIQUE(_xlfn.HSTACK(_xlfn.ANCHORARRAY('Contract Details'!B12)))),1,1,0)))\r\n), 2,1,FALSE),\r\n_xlfn.HSTACK(name_CompanyCode,\"[All Contracts]\",name_Account_GL_Current,\"Current\"),\r\n_xlfn.HSTACK(name_CompanyCode,\"[All Contracts]\",name_Account_GL_NonCurrent,\"Non-Current\"),\r\n_xlfn.HSTACK(name_CompanyCode,\"[All Contracts]\",name_Account_GL_Total,\"🔼 Total\")\r\n))";
                ws.Cells["A12"].Calculate();

                List<string> cellValuesAfter = new();
                int i = 0;
                foreach (var cell in ws.Cells["D12:D100"])
                {
                    Assert.AreEqual(cell.Text, cellValues[i]);
                    i++;
                }

                ws.ClearFormulas();
                SaveAndCleanup(pck);
            }
        }
        public void s871()
        {
            using var epplusCalculated = OpenTemplatePackage("s871.xlsx");
            using var excelCalculated = OpenTemplatePackage("s871.xlsx");

            ExcelPackage.MemorySettings.UseRecyclableMemory = false;

            Console.WriteLine("Calculating...");

            var wbEPPlus = epplusCalculated.Workbook;
            wbEPPlus.CalcMode = ExcelCalcMode.Manual;

            epplusCalculated.Workbook.Calculate(new OfficeOpenXml.FormulaParsing.ExcelCalculationOption
            {
                PrecisionAndRoundingStrategy = OfficeOpenXml.FormulaParsing.PrecisionAndRoundingStrategy.Excel,
                AllowCircularReferences = true,
                EnableUnicodeAwareStringOperations = true,
            });

            var trackCol = 7;
            var trackRow = 12;

            var sheet = wbEPPlus.Worksheets["Summary"];

            object o;

            o = sheet.Cells[trackRow, trackCol].Formula;

            Assert.AreEqual(6384484.31, (double)sheet.Cells["G9"].Value);
            Assert.AreEqual(174125.25, (double)sheet.Cells["G10"].Value);
            Assert.AreEqual(71467.24, (double)sheet.Cells["G11"].Value);
            Assert.AreEqual(32018.5, (double)sheet.Cells["G12"].Value);
        }
        [TestMethod]
        public void s871_2()
        {
            using var p = OpenTemplatePackage("s871-2.xlsx");

            Console.WriteLine("Calculating...");

            var wbEPPlus = p.Workbook;

            p.Workbook.Calculate(new ExcelCalculationOption
            {
                PrecisionAndRoundingStrategy = OfficeOpenXml.FormulaParsing.PrecisionAndRoundingStrategy.Excel,
                AllowCircularReferences = true,
                EnableUnicodeAwareStringOperations = true,
            });


            var sheet1 = wbEPPlus.Worksheets["FTE_BSTM"];
            var sheet2 = wbEPPlus.Worksheets["Aico Data"];

            Assert.AreEqual("CHF", sheet1.Cells["I51"].Value);
            Assert.AreEqual("CHF", sheet1.Cells["I52"].Value);
            Assert.AreEqual("0280", sheet2.Cells["A77"].Value);
            Assert.AreEqual("CHF", sheet2.Cells["H77"].Value);
        }
        [TestMethod]

        public void s875()
        {
            using (var p = OpenTemplatePackage("s875.xlsx"))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets["List of Orders"];
                wb.Calculate();
                Assert.AreEqual("", ws.Cells["A13"].Value);
                Assert.AreEqual("Total", ws.Cells["F13"].Value);
                Assert.AreEqual("[varioius]", ws.Cells["G13"].Value);
                Assert.AreEqual(5505974.18, (double)ws.Cells["N13"].Value, 0.0001);
                Assert.AreEqual(-4101976.8, (double)ws.Cells["O13"].Value, 0.0001);
                Assert.AreEqual(-1083371.20, (double)ws.Cells["P13"].Value, 0.0001);
            }
        }
        //-------

        [TestMethod]
        public void LeftHas_One_MinimumParameter()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("AnchorDynamic");

                ws.Cells["B3"].Value = "Hello World!";
                //Left with num_chars omitted takes 1 char.
                ws.Cells["C3"].Formula = "LEFT(B3)";

                ws.Calculate();

                Assert.AreEqual("H", ws.Cells["C3"].Value);
            }
        }

        //Part of s884
        [TestMethod]
        public void AnchorArray_DynamicArrayFormula_Single()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("AnchorDynamic");

                //Set formulas that will produce value
                ws.Cells["E4"].Formula = "SUM(2+2)";
                ws.Cells["G4"].Formula = "HSTACK($E$4)";

                //Set AnchorArray on single cells
                ws.Cells["C3"].Formula = "ANCHORARRAY(E4)";
                ws.Cells["C4"].Formula = "ANCHORARRAY(G4)";

                //Single Cell Formulas with AnchorArray should return value after lookup, not #REF!
                ws.Calculate();
                Assert.AreEqual(4d, ws.Cells["C3"].Value);
                Assert.AreEqual(4d, ws.Cells["C4"].Value);
            }
        }

        //Part of s884
        [TestMethod]
        public void XlookupSingleValue()
        {
            using (var p = OpenPackage("TestXLookupSingle.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("xLookup");

                //Create a lookup range of values
                var lookupRange = ws.Cells["F2:G6"];
                for (int i = 0; i < lookupRange.Rows; i++)
                {
                    for (int j = 0; j < lookupRange.Columns; j++)
                    {
                        lookupRange.SetCellValue(i, j, j == 0 ? i : i * 10);
                    }
                }

                //Create spillover formula with only one value
                ws.Cells["P2"].Value = 4;
                ws.Cells["Q2"].Formula = "HSTACK($P$2)";
                //Access it via AnchorArray and  run XLookup;
                ws.Cells["A1"].Formula = "XLOOKUP(ANCHORARRAY($Q$2),F1:F6,G1:G6,\"fallbackValue\",0,1)";

                ws.Calculate();

                Assert.AreEqual(40, ws.Cells["A1"].Value);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void s884()
        {
            using (var p = OpenTemplatePackage("s884.xlsx"))
            {
                var ws = p.Workbook.Worksheets["Summary"];
                var technical = p.Workbook.Worksheets["Technical"];

                var form = ws.Cells["A12"].Formula;
                var formSimpler = technical.Cells["K19"].Formula;

                ws.Calculate(o => o.EnableUnicodeAwareStringOperations = true);

                var someVal = ws.Cells["E12"].Value;
                var lcTest = ws.Cells["F12"].Value;
                var amountLC = ws.Cells["G12"].Value;

                Assert.AreEqual(ws.Cells["A12"].Text, "0110");
                Assert.AreEqual(ws.Cells["B12"].Text, "200150");
                Assert.AreEqual(ws.Cells["C12"].Text, "Sub-Total for EUR");
                Assert.AreEqual(ws.Cells["D12"].Text, "EUR");
                Assert.AreEqual(-200000d, ws.Cells["E12"].Value);
                Assert.AreEqual(ws.Cells["F12"].Text, "CHF");
                Assert.AreEqual(-200000d, ws.Cells["G12"].Value);

                var dataSheet = p.Workbook.Worksheets[9];

                dataSheet.Calculate(o => o.EnableUnicodeAwareStringOperations = true);

                Assert.AreEqual(dataSheet.Cells["D31"].Text, "0280");
                Assert.AreEqual(dataSheet.Cells["AG31"].Text, "0110_vs_0280_Loans_Current_Received_EUR");
                Assert.AreEqual(dataSheet.Cells["AH31"].Text, "0110_vs_0280_Ref_");
            }
        }

        [TestMethod]
        public void s868()
        {
            using (var package = OpenTemplatePackage("s868.xlsx"))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets["aliss"];
                wb.Worksheets.Add("AlissCopy", ws);
                ws.Cells["CW151"].Calculate();
                Assert.AreEqual(39.29, ws.Cells["CW151"].Value);
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void Copy_Names_to_new_workbook()
        {
            using (ExcelPackage origPck = OpenPackage("i345_ORIG_Global_Names.xlsx", true))
            {
                var wbOrig = origPck.Workbook;
                var wsOrig = wbOrig.Worksheets.Add("Testpage");

                wbOrig.Names.AddFormula("MyDefinedFormula", "Testpage!C3");//This line fails
                wbOrig.Names.AddValue("MyDefinedValue", 10);
                var range = wbOrig.Names.Add("MyRange", wsOrig.Cells["C3:D4"]);

                wsOrig.Cells["C3"].Formula = "5+5";
                wsOrig.Cells["D4"].Formula = "MyDefinedFormula";
                wsOrig.Cells["G10"].Formula = "MyRange";

                wsOrig.Calculate();

                using (ExcelPackage destPackage = OpenPackage("i345_DEST_Global_Names.xlsx", true))
                {
                    var target = destPackage.Workbook.Worksheets.Add("destWs", wsOrig);
                    target.Calculate();
                    SaveAndCleanup(destPackage);
                }
                SaveAndCleanup(origPck);
            }
        }
        [TestMethod]
        public void TestNames()
        {
            using (var origPck = OpenPackage("copyworkbooknames.xlsx", true))
            {
                var wbOrig = origPck.Workbook;
                var wsOrig = wbOrig.Worksheets.Add("Testpage");

                wbOrig.Names.AddFormula("MyDefinedFormula", "Testpage!C3");//This line fails
                wbOrig.Names.AddValue("MyDefinedValue", 10);
                var range = wbOrig.Names.Add("MyRange", wsOrig.Cells["C3:D4"]);

                wsOrig.Cells["C3"].Formula = "5+5";
                wsOrig.Cells["D4"].Formula = "MyDefinedFormula";
                wsOrig.Cells["G10"].Formula = "MyRange";

                wsOrig.Calculate();
                origPck.Save();

                using (ExcelPackage destPackage = new ExcelPackage(origPck.Stream))
                {
                    var target = destPackage.Workbook.Worksheets.Add("destWs", wsOrig);
                    target.Calculate();
                    SaveAndCleanup(destPackage);
                }
                SaveAndCleanup(origPck);
            }
        }
        [TestMethod]
        public void s927()
        {
            using var p = OpenTemplatePackage("s927.xlsx");
            //using var p = OpenTemplatePackage("s927 - Calced.xlsx");
            var ws = p.Workbook.Worksheets["Calculation sheet"];
            ws.Calculate();
            Assert.AreEqual(938643.13, ws.Cells["H10"].Value);
        }
        [TestMethod]
        public void s965_1()
        {
            using var p = OpenTemplatePackage("Aico\\s965-1.xlsx");
            var ws = p.Workbook.Worksheets["Aico Data"];
            ws.Calculate();
            Assert.AreEqual(-64440.8652, (double)ws.Cells["D41"].Value, 0.0001);
            Assert.AreEqual(-3206585.8006, (double)ws.Cells["D42"].Value, 0.0001);
        }
        [TestMethod]
        public void s965_2()
        {
            using var p = OpenTemplatePackage("Aico\\s965-2.xlsx");
            var ws = p.Workbook.Worksheets["Aico Data"];
            ws.Calculate();
            Assert.AreEqual("", ws.Cells["B49"].Value);
            Assert.AreEqual("7300030", ws.Cells["B50"].Value);
            ws = p.Workbook.Worksheets["Calculation"];
            Assert.AreEqual(4D, ws.Cells["AP2"].Value);
            Assert.AreEqual(9D, ws.Cells["AQ2"].Value);
        }
        [TestMethod]
        public void s968()
        {
            using var p = OpenTemplatePackage("s968.xlsx");
            var ws = p.Workbook.Worksheets["Messages"];

            ws.Cells["D17"].Calculate();

            Assert.AreEqual("Error", ws.Cells["D18"].Value);
            Assert.AreEqual("Warning", ws.Cells["D26"].Value);
            Assert.AreEqual("Error", ws.Cells["D40"].Value);
            Assert.AreEqual("One or more errors occurred during the valuation run", ws.Cells["E40"].Value);
            Assert.IsNull(ws.Cells["D41"].Value);
        }
    }
}

