using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class DefinedNameIssues : TestBase
    {
		[TestMethod]
		public void s652()
        {
            using (var p = OpenTemplatePackage("s652.xlsm"))
            {
                using var p2 = new ExcelPackage();
                var ws = p.Workbook.Worksheets[0];
                p2.Workbook.Worksheets.Add("New ws", ws);
                SaveWorkbook("s652.xlsx", p2);
            }
		}
        [TestMethod]
        //i1408
        public void VersionName()
        {
            using (var package = OpenTemplatePackage("VersionNameManager.xlsx"))
            {
                var name = package.Workbook.Names.First();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        //i1408
        public void DefinedNamesQuoteError()
        {
            using (var package = OpenPackage("QuoteError.xlsx", true))
            {
                package.Workbook.Worksheets.Add("something");
                package.Workbook.Names.AddValue("Lae_Zel", "zhak vo\"n\"fynh duj");

                var packageTemp = OpenPackage("dummyQuoteWorkbook.xlsx", true);
                packageTemp.Workbook.Worksheets.Add("dummy");
                SaveAndCleanup (packageTemp);

                var file = new FileInfo("C:\\epplusTest\\Testoutput\\dummyQuoteWorkbook.xlsx");

                package.Workbook.ExternalLinks.AddExternalWorkbook(file);

                package.Workbook.Names.AddFormula("编制单位", "\"编制单位：\"&[1]dummyQuoteWorkbook!$D$6");


                package.Workbook.Names.AddValue("Unended", "s\"omething");
                package.Workbook.Names.AddValue("EndedRepeated", "s\"\"omething");

                SaveAndCleanup(package);
            }

            using (var package = OpenPackage("QuoteError.xlsx"))
            {

                Assert.AreEqual("zhak vo\"n\"fynh duj", package.Workbook.Names["Lae_Zel"].Value);
                Assert.AreEqual("\"编制单位：\"&[1]dummyQuoteWorkbook!$D$6", package.Workbook.Names["编制单位"].Formula);
                Assert.AreEqual("s\"omething", package.Workbook.Names["Unended"].Value);
                Assert.AreEqual("s\"\"omething", package.Workbook.Names["EndedRepeated"].Value);


                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void issue2224()
        {

            var cases = new (string Name, string Description, Action<ExcelWorkbook, ExcelWorksheet> Setup)[]
            {
                (
                    "SciArrayFormula",
                    "Inline array literal containing scientific-notation constants stored in Formula",
                    (workbook, sheet) =>
                    {
                        workbook.Names.AddFormula(
                            "SciArrayFormula",
                            "{4.02506300418233E-305,3.33761291040418E-308}");
                        sheet.Cells["B1"].Formula = "SciArrayFormula";
                    }
                ),
                (
                    "UndefinedUdfName",
                    "Workbook-level name that references an undefined UDF (Main.SAPF4Help)",
                    (workbook, sheet) =>
                    {
                        workbook.Names.AddFormula("SAPFuncF4Help", "Main.SAPF4Help()");
                    }
                ),
                (
                    "CubeSetName",
                    "Workbook-level name that uses CUBESET against ThisWorkbookDataModel",
                    (workbook, sheet) =>
                    {
                        workbook.Names.AddFormula(
                            "Slicer_PC_P210",
                            "CUBESET(\"ThisWorkbookDataModel\",\"[DIM_PC].[PC_P2].&[RS]\",\"Slicer\")");
                    }
                )
            };
            var i = 1;
            foreach (var (name, description, setup) in cases)
            {
                var xlsxFile = $"issue2224-{i++}.xlsx";
                using (var package = OpenPackage(xlsxFile, true))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells["A1"].Value = 1;
                    setup(package.Workbook, worksheet);
                    SaveAndCleanup(package);
                }

                using var reopened = new ExcelPackage(new FileInfo(xlsxFile));

                Console.WriteLine($"Case: {name}");
                Console.WriteLine($"  Description: {description}");

                try
                {
                    reopened.Workbook.Calculate(new ExcelCalculationOption { AllowCircularReferences = true });
                    Console.WriteLine("  Result: calculation succeeded (unexpected)\n");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Result: {ex.GetType().Name} - {ex.Message}\n");
                }
            }
        }
        [TestMethod]
        public void issue2226()
        {
            static (ExcelPackage pkg, ExcelWorksheet ws1, ExcelWorksheet ws2) CreateWorkbook(bool includeSheetScopedName)
            {
                var pkg = new ExcelPackage();
                var ws1 = pkg.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pkg.Workbook.Worksheets.Add("Sheet2");

                // Workbook-scoped name "MyTable" (should be used by formulas on Sheet2)
                ws2.Cells["A1"].Value = 1;
                ws2.Cells["B1"].Value = 2;
                ws2.Cells["A2"].Value = 10;
                ws2.Cells["B2"].Value = 20;
                pkg.Workbook.Names.Add("MyTable", ws2.Cells["A1:B2"]);

                if (includeSheetScopedName)
                {
                    // Sheet-scoped name "MyTable" on Sheet1 (should NOT affect formulas on Sheet2)
                    ws1.Cells["A1"].Value = 1;
                    ws1.Cells["B1"].Value = 2;
                    ws1.Cells["A2"].Value = 1;
                    ws1.Cells["B2"].Value = 2;
                    ws1.Names.Add("MyTable", ws1.Cells["A1:B2"]);
                }

                // Put the same formula in multiple cells so we can test different calculation APIs
                // without accidental caching/order effects between tests.
                ws2.Cells["C1"].Formula = "HLOOKUP(1,MyTable,2,FALSE)"; // used for string-eval + address-eval
                ws2.Cells["C2"].Formula = "HLOOKUP(1,MyTable,2,FALSE)"; // used for range.Calculate()
                ws2.Cells["C3"].Formula = "HLOOKUP(1,MyTable,2,FALSE)"; // used for workbook.Calculate()
                return (pkg, ws1, ws2);
            }

            static void RunTest(string name, Func<(ExcelPackage pkg, ExcelWorksheet ws1, ExcelWorksheet ws2), string> run)
            {
                Console.WriteLine($"\n=== {name} ===");
                var ctx = CreateWorkbook(includeSheetScopedName: true);
                using (ctx.pkg)
                {
                    Console.WriteLine(run(ctx));
                }
            }

            Console.WriteLine("Expected (Excel semantics): 10");

            RunTest("Mode A: ws.Calculate(formula-string) is wrong", ctx =>
            {
                object? inWs2;
                object? inWs1;
                try { inWs2 = ctx.ws2.Calculate(ctx.ws2.Cells["C1"].Formula); }
                catch (Exception ex) { inWs2 = $"EXCEPTION: {ex.GetType().Name}: {ex.Message}"; }
                Assert.AreEqual(inWs2, 10);
                try { inWs1 = ctx.ws1.Calculate(ctx.ws2.Cells["C1"].Formula); }
                catch (Exception ex) { inWs1 = $"EXCEPTION: {ex.GetType().Name}: {ex.Message}"; }
                Assert.AreEqual(inWs1, 1);
                return $"ws2.Calculate(formula) => {inWs2}\nws1.Calculate(formula) => {inWs1}";
            });

            RunTest("Mode B2: range.Calculate() is right", ctx =>
            {
                var before = ctx.ws2.Cells["C2"].Value;
                try { ctx.ws2.Cells["C2"].Calculate(); }
                catch (Exception ex) { return $"Before => {before}\nEXCEPTION: {ex.GetType().Name}: {ex.Message}"; }
                Assert.AreEqual(ctx.ws2.Cells["C2"].Value, 10);
                return $"Before => {before}\nAfter  => {ctx.ws2.Cells["C2"].Value}";
            });

            RunTest("Mode B3: worksheet.Calculate() is right", ctx =>
            {
                var before = ctx.ws2.Cells["C2"].Value;
                try { ctx.ws2.Calculate(); }
                catch (Exception ex) { return $"Before => {before}\nEXCEPTION: {ex.GetType().Name}: {ex.Message}"; }
                Assert.AreEqual(ctx.ws2.Cells["C2"].Value, 10);
                return $"Before => {before}\nAfter  => {ctx.ws2.Cells["C2"].Value}";
            });

            RunTest("Mode B: workbook.Calculate() is right", ctx =>
            {
                var before = ctx.ws2.Cells["C3"].Value;
                try { ctx.pkg.Workbook.Calculate(); }
                catch (Exception ex) { return $"Before => {before}\nEXCEPTION: {ex.GetType().Name}: {ex.Message}"; }
                Assert.AreEqual(ctx.ws2.Cells["C2"].Value, 10);
                return $"Before => {before}\nAfter  => {ctx.ws2.Cells["C3"].Value}";
            });

            RunTest("Mode C: ws.Calculate(address) is right", ctx =>
            {
                object? fromWs2;
                object? fromWs1;
                try { fromWs2 = ctx.ws2.Calculate("'Sheet2'!C1"); }
                catch (Exception ex) { fromWs2 = $"EXCEPTION: {ex.GetType().Name}: {ex.Message}"; }
                Assert.AreEqual(fromWs2, 10);
                try { fromWs1 = ctx.ws1.Calculate("'Sheet2'!C1"); }
                catch (Exception ex) { fromWs1 = $"EXCEPTION: {ex.GetType().Name}: {ex.Message}"; }
                return $"ws2.Calculate(\"'Sheet2'!C1\") => {fromWs2}\nws1.Calculate(\"'Sheet2'!C1\") => {fromWs1}";
                //Assert.AreEqual(fromWs1, 10);
            });

            RunTest("Sanity: removing sheet-scoped name fixes formula-string eval", ctx =>
            {
                // Demonstrate the fix within one workbook instance.
                object? before;
                object? after;
                try { before = ctx.ws2.Calculate(ctx.ws2.Cells["C1"].Formula); }
                catch (Exception ex) { before = $"EXCEPTION: {ex.GetType().Name}: {ex.Message}"; }

                ctx.ws1.Names.Remove("MyTable");

                try { after = ctx.ws2.Calculate(ctx.ws2.Cells["C1"].Formula); }
                catch (Exception ex) { after = $"EXCEPTION: {ex.GetType().Name}: {ex.Message}"; }

                return $"Before removing ws1-scoped name => {before}\nAfter  removing ws1-scoped name => {after}";
            });
        }
    }
}
