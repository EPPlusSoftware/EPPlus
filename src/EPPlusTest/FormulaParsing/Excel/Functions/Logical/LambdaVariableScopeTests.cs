using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.DependencyChain;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    /// <summary>
    /// Tests for variable scoping in LET/LAMBDA, see GitHub issue #2606.
    /// A named LAMBDA with a LET body must resolve its parameters lexically,
    /// never from the scope of another LAMBDA call that happens to be on the variable storage stack.
    /// </summary>
    [TestClass]
    public class LambdaVariableScopeTests
    {
        #region Issue 2606 - repro 2 (single workbook calculation)

        // Expected to FAIL before fix (returns 56: Out's LET body reads ID_in = 5 from In(5)).
        [TestMethod]
        public void NamedLambdaWithLet_NestedNamedLambdaArgument_SameParameterName()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("A");
            wb.Names.AddFormula("Out", "LAMBDA(ID_in, X, LET(v, ID_in*10, v+X))");
            wb.Names.AddFormula("In", "LAMBDA(ID_in, ID_in+1)");
            sheet.Cells["A1"].Formula = "Out(2, In(5))";
            wb.Calculate();
            Assert.AreEqual(26d, sheet.Cells["A1"].Value);
        }

        // Expected to FAIL before fix. B1 is calculated while Out's arguments are processed.
        [TestMethod]
        public void NamedLambdaWithLet_ArgumentFromCellCalculatedDuringCall_SameParameterName()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("A");
            wb.Names.AddFormula("Out", "LAMBDA(ID_in, X, LET(v, ID_in*10, v+X))");
            wb.Names.AddFormula("In", "LAMBDA(ID_in, ID_in+1)");
            sheet.Cells["A1"].Formula = "Out(2, B1)";
            sheet.Cells["B1"].Formula = "In(5)";
            wb.Calculate();
            Assert.AreEqual(6d, sheet.Cells["B1"].Value);
            Assert.AreEqual(26d, sheet.Cells["A1"].Value);
        }

        // Control: expected to PASS before and after fix (no LET in Out).
        [TestMethod]
        public void NamedLambdaWithoutLet_NestedNamedLambdaArgument_SameParameterName()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("A");
            wb.Names.AddFormula("Out", "LAMBDA(ID_in, X, ID_in*10+X)");
            wb.Names.AddFormula("In", "LAMBDA(ID_in, ID_in+1)");
            sheet.Cells["A1"].Formula = "Out(2, In(5))";
            wb.Calculate();
            Assert.AreEqual(26d, sheet.Cells["A1"].Value);
        }

        // Control: expected to PASS before and after fix (different parameter names).
        [TestMethod]
        public void NamedLambdaWithLet_NestedNamedLambdaArgument_DifferentParameterName()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("A");
            wb.Names.AddFormula("Out", "LAMBDA(Dout, X, LET(v, Dout*10, v+X))");
            wb.Names.AddFormula("In", "LAMBDA(ID_in, ID_in+1)");
            sheet.Cells["A1"].Formula = "Out(2, In(5))";
            wb.Calculate();
            Assert.AreEqual(26d, sheet.Cells["A1"].Value);
        }

        // Control: expected to PASS before and after fix (precedent calculated before the calling cell).
        [TestMethod]
        public void NamedLambdaWithLet_ArgumentFromCellCalculatedBeforeCall_SameParameterName()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("A");
            wb.Names.AddFormula("Out", "LAMBDA(ID_in, X, LET(v, ID_in*10, v+X))");
            wb.Names.AddFormula("In", "LAMBDA(ID_in, ID_in+1)");
            sheet.Cells["A1"].Formula = "In(5)";
            sheet.Cells["A2"].Formula = "Out(2, A1)";
            wb.Calculate();
            Assert.AreEqual(6d, sheet.Cells["A1"].Value);
            Assert.AreEqual(26d, sheet.Cells["A2"].Value);
        }

        #endregion

        #region Issue 2606 - repro 1 (worksheet by worksheet calculation)

        private static ExcelPackage CreateRepro1Package()
        {
            var package = new ExcelPackage();
            var wb = package.Workbook;
            var a = wb.Worksheets.Add("A");
            var b = wb.Worksheets.Add("B");
            var c = wb.Worksheets.Add("C");
            wb.Names.AddFormula("Kf", "LAMBDA(t, 2*t)");
            wb.Names.AddFormula("Loss", "LAMBDA(Dh_in, K, LET(v, 1/Dh_in, v*K))");
            wb.Names.AddFormula("Slip", "LAMBDA(K, Dh_in, K+Dh_in)");
            a.Cells["A1"].Formula = "Kf(5)";
            wb.Names.Add("K", a.Cells["A1"]);
            a.Cells["A2"].Formula = "Loss(4,K)";
            b.Cells["A1"].Formula = "Slip(K,8)";
            c.Cells["A1"].Formula = "A!A2";
            return package;
        }

        // Expected to FAIL before fix (A!A2 returns 1.25: Loss reads Dh_in = 8 from Slip(K,8)).
        [TestMethod]
        public void NamedLambdaWithLet_CalculatePerWorksheet_SameParameterNameInOtherLambda()
        {
            using var package = CreateRepro1Package();
            var wb = package.Workbook;
            foreach (var ws in wb.Worksheets)
            {
                ws.Calculate();
            }
            Assert.AreEqual(10d, wb.Worksheets["A"].Cells["A1"].Value);
            Assert.AreEqual(2.5d, wb.Worksheets["A"].Cells["A2"].Value);
            Assert.AreEqual(18d, wb.Worksheets["B"].Cells["A1"].Value);
            Assert.AreEqual(2.5d, wb.Worksheets["C"].Cells["A1"].Value);
        }

        // Control: expected to PASS before and after fix.
        [TestMethod]
        public void NamedLambdaWithLet_CalculateWorkbook_SameParameterNameInOtherLambda()
        {
            using var package = CreateRepro1Package();
            var wb = package.Workbook;
            wb.Calculate();
            Assert.AreEqual(10d, wb.Worksheets["A"].Cells["A1"].Value);
            Assert.AreEqual(2.5d, wb.Worksheets["A"].Cells["A2"].Value);
            Assert.AreEqual(18d, wb.Worksheets["B"].Cells["A1"].Value);
            Assert.AreEqual(2.5d, wb.Worksheets["C"].Cells["A1"].Value);
        }

        #endregion

        #region Related - caller scope must not leak into a called named lambda

        // Expected value to be verified in Excel. A named lambda is defined at workbook level and
        // should not see the caller's parameters, so X in G is expected to be unresolved (#NAME?).
        [TestMethod]
        public void NamedLambda_ShouldNotSeeCallerNamedLambdaParameters()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("Sheet1");
            wb.Names.AddFormula("G", "LAMBDA(a, a+X)");
            wb.Names.AddFormula("H", "LAMBDA(X, G(1))");
            sheet.Cells["A1"].Formula = "H(5)";
            wb.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells["A1"].Value);
        }

        // Expected value to be verified in Excel. Same as above but the caller is an inline lambda.
        [TestMethod]
        public void NamedLambda_ShouldNotSeeCallerInlineLambdaParameters()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("Sheet1");
            wb.Names.AddFormula("G", "LAMBDA(a, a+X)");
            sheet.Cells["A1"].Formula = "LAMBDA(X, G(1))(5)";
            wb.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells["A1"].Value);
        }

        #endregion

        #region Related - named lambda called with too few arguments

        // Expected value to be verified in Excel. Currently probably passes via an
        // ArgumentOutOfRangeException in LambdaNameFunction.Execute that is caught and turned into #VALUE!.
        [TestMethod]
        public void NamedLambda_TooFewArguments_ShouldReturnValueError()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("Sheet1");
            wb.Names.AddFormula("Out", "LAMBDA(ID_in, X, LET(v, ID_in*10, v+X))");
            sheet.Cells["A1"].Formula = "Out(2)";
            wb.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A1"].Value);
        }

        #endregion

        #region Step 2 - variable storage stack invariants

        private static int GetVariableStorageCount(RpnOptimizedDependencyChain depChain)
        {
            var storage = depChain._parsingContext.VariableStorage;
            return storage == null ? 0 : storage.Count;
        }

        private static void AssertCompilePreservesOuterStack(string formula)
        {
            var tokens = SourceCodeTokenizer.Default.Tokenize(formula);
            var rpnTokens = FormulaExecutor.CreateRPNTokens(tokens);
            var ctx = ParsingContext.Create();
            ctx.VariableStorage = new VariableStorageManager();
            var outerScope = ctx.VariableStorage.AddNewScope();
            LambdaFormulaSettings lambdaSettings = default;

            FormulaExecutor.CompileExpressions(ref lambdaSettings, ref rpnTokens, ctx);

            Assert.AreEqual(1, ctx.VariableStorage.Count, "Compilation must restore the stack depth, not clear it.");
            Assert.AreSame(outerScope, ctx.VariableStorage.Peek());
        }

        // Expected to FAIL before step 2 (CompileExpressions calls Clear()).
        [TestMethod]
        public void CompileLet_ShouldPreserveOuterVariableStorageStack()
        {
            AssertCompilePreservesOuterStack("LET(x,1,x+1)");
        }

        // Expected to FAIL before step 2 (CompileExpressions calls Clear()).
        [TestMethod]
        public void CompileInlineLambda_ShouldPreserveOuterVariableStorageStack()
        {
            AssertCompilePreservesOuterStack("LAMBDA(a,a+1)(1)");
        }

        // Control: expected to PASS before and after step 2 (LET pops its scope in OnDispose).
        [TestMethod]
        public void VariableStorage_ShouldBeEmptyAfterCalculate_Let()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LET(x,2,y,LET(a,x,a+1),x + y)";
            var depChain = RpnFormulaExecution.Execute(package.Workbook, new ExcelCalculationOption());
            Assert.AreEqual(5d, sheet.Cells["A1"].Value);
            Assert.AreEqual(0, GetVariableStorageCount(depChain));
        }

        // Expected to FAIL before step 2 (LambdaInvoker never calls LambdaInvokeExpression.OnDispose).
        [TestMethod]
        public void VariableStorage_ShouldBeEmptyAfterCalculate_InlineLambda()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(x,x+1)(1)";
            var depChain = RpnFormulaExecution.Execute(package.Workbook, new ExcelCalculationOption());
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            Assert.AreEqual(0, GetVariableStorageCount(depChain));
        }

        // Expected to FAIL before step 2 (scope pushed in LambdaCalculator.Execute is never popped for named lambdas).
        [TestMethod]
        public void VariableStorage_ShouldBeEmptyAfterCalculate_NestedNamedLambdas()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("A");
            wb.Names.AddFormula("Out", "LAMBDA(ID_in, X, LET(v, ID_in*10, v+X))");
            wb.Names.AddFormula("In", "LAMBDA(ID_in, ID_in+1)");
            sheet.Cells["A1"].Formula = "Out(2, In(5))";
            var depChain = RpnFormulaExecution.Execute(wb, new ExcelCalculationOption());
            Assert.AreEqual(26d, sheet.Cells["A1"].Value);
            Assert.AreEqual(0, GetVariableStorageCount(depChain));
        }

        // Status before step 2 unknown.
        [TestMethod]
        public void VariableStorage_ShouldBeEmptyAfterCalculate_RecursiveNamedLambda()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Names.AddFormula("Factorial2", "LAMBDA(n, IF(n = 0, 1, n * Sheet1!Factorial2(n - 1)))");
            sheet.Cells["A1"].Formula = "Factorial2(4)";
            var depChain = RpnFormulaExecution.Execute(package.Workbook, new ExcelCalculationOption());
            Assert.AreEqual(24d, sheet.Cells["A1"].Value);
            Assert.AreEqual(0, GetVariableStorageCount(depChain));
        }

        // Status before step 2 unknown.
        [TestMethod]
        public void VariableStorage_ShouldBeEmptyAfterCalculate_LambdaHelperFunctions()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "REDUCE(0, {1,2,3}, LAMBDA(a,v,a+v))";
            sheet.Cells["B1"].Formula = "SUM(MAP({1,2,3}, LAMBDA(x, LET(y, x*2, y))))";
            var depChain = RpnFormulaExecution.Execute(package.Workbook, new ExcelCalculationOption());
            Assert.AreEqual(6d, sheet.Cells["A1"].Value);
            Assert.AreEqual(12d, sheet.Cells["B1"].Value);
            Assert.AreEqual(0, GetVariableStorageCount(depChain));
        }

        // Expected to FAIL before step 2 (scopes left from one worksheet calculation remain for the next).
        [TestMethod]
        public void VariableStorage_ShouldBeEmptyAfterEachWorksheetCalculate()
        {
            using var package = CreateRepro1Package();
            foreach (var ws in package.Workbook.Worksheets)
            {
                var depChain = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
                Assert.AreEqual(0, GetVariableStorageCount(depChain), $"Variable storage not empty after calculating worksheet {ws.Name}");
            }
            Assert.AreEqual(2.5d, package.Workbook.Worksheets["A"].Cells["A2"].Value);
        }

        #endregion

        [TestMethod]
        public void NamedLambda_TooManyArguments_ShouldReturnValueError()
        {
            using var package = new ExcelPackage();
            var wb = package.Workbook;
            var sheet = wb.Worksheets.Add("Sheet1");
            wb.Names.AddFormula("Out", "LAMBDA(ID_in, X, LET(v, ID_in*10, v+X))");
            sheet.Cells["A1"].Formula = "Out(2, 3, 4)";
            wb.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A1"].Value);
        }
    }
}