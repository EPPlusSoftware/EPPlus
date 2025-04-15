using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.DependencyChain;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    /*
     * Status 2025-02-21: The functionality in the tokenizer/RPN tokens that concatenated comma-separated
     * Excel addresses is disabled. We need unit tests for this, see line 76 in FormulaExecutor.CreateRpnTokens.
     */
    [TestClass]
    public class LambdaFunctionTests
    {
        [TestMethod]
        public void LambdaSelfInvokeSingleArg()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(r,r+1)(D6)";
            sheet.Cells["D6"].Value = 5;
            sheet.Calculate();
            Assert.AreEqual(6d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void LambdaSelfInvokeTest1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(r,c,r+c)(D6,D7)";
            sheet.Cells["D6"].Value = 5;
            sheet.Cells["D7"].Value = 6;
            sheet.Calculate();
            Assert.AreEqual(11d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void LambdaTokensTest1()
        {
            var tokens = SourceCodeTokenizer.Default.Tokenize("LAMBDA(r,c,r+c)(D6,D7)");
            var rpnTokens = FormulaExecutor.CreateRPNTokens(tokens);
            Assert.AreEqual(1, rpnTokens.LambdaRefs.Count);
            Assert.AreEqual(0, rpnTokens.LambdaRefs.First().Key);
            Assert.AreEqual(9, rpnTokens.LambdaRefs.First().Value);

            var ctx = ParsingContext.Create();
            var lambdaSettings = default(LambdaFormulaSettings);
            var exp = FormulaExecutor.CompileExpressions(ref lambdaSettings, ref rpnTokens, ctx);
            Assert.AreEqual(6, exp.Count);
            Assert.AreEqual(ExpressionType.Function, exp[0].ExpressionType);
            Assert.IsInstanceOfType(exp[1], typeof(VariableExpression));
            Assert.IsInstanceOfType(exp[3], typeof(VariableExpression));
            Assert.AreEqual(ExpressionType.LambdaCalculation, exp[4].ExpressionType);
            Assert.IsInstanceOfType(exp[4], typeof(LambdaTokensExpression));
            Assert.IsInstanceOfType(exp[9], typeof(RangeExpression));
            Assert.IsInstanceOfType(exp[11], typeof(RangeExpression));
        }

        [TestMethod]
        public void LambdaTokensTest2()
        {
            var tokens = SourceCodeTokenizer.Default.Tokenize("LAMBDA(r,c,r+c)(D6,D7)");
            var rpnTokens = FormulaExecutor.CreateRPNTokens(tokens);
            Assert.AreEqual(1, rpnTokens.LambdaRefs.Count);
            Assert.AreEqual(0, rpnTokens.LambdaRefs.First().Key);
            Assert.AreEqual(9, rpnTokens.LambdaRefs.First().Value);

            var ctx = ParsingContext.Create();
            var lambdaSettings = default(LambdaFormulaSettings);
            var exp = FormulaExecutor.CompileExpressions(ref lambdaSettings, ref rpnTokens, ctx);
            Assert.AreEqual(6, exp.Count);
            Assert.AreEqual(ExpressionType.Function, exp[0].ExpressionType);
            Assert.IsInstanceOfType(exp[1], typeof(VariableExpression));
            Assert.IsInstanceOfType(exp[3], typeof(VariableExpression));
            Assert.AreEqual(ExpressionType.LambdaCalculation, exp[4].ExpressionType);
            Assert.IsInstanceOfType(exp[4], typeof(LambdaTokensExpression));
            Assert.IsInstanceOfType(exp[9], typeof(RangeExpression));
            Assert.IsInstanceOfType(exp[11], typeof(RangeExpression));
        }

        [TestMethod]
        public void LambdaTokensTest3()
        {
            var tokens = SourceCodeTokenizer.Default.Tokenize("LAMBDA(a, a + LAMBDA(b, b + a)(2))(2)");
            var rpnTokens = FormulaExecutor.CreateRPNTokens(tokens);
            Assert.AreEqual(2, rpnTokens.LambdaRefs.Count);
            Assert.AreEqual(4, rpnTokens.LambdaRefs.First().Key);
            Assert.AreEqual(11, rpnTokens.LambdaRefs.First().Value);

            var ctx = ParsingContext.Create();
            LambdaFormulaSettings lambdaSettings = default;
            var exp = FormulaExecutor.CompileExpressions(ref lambdaSettings, ref rpnTokens, ctx);
            Assert.AreEqual(4, exp.Count);
        }

        [TestMethod]
        public void LambdaSelfInvokeTest2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IF(TRUE(),LAMBDA(r,c,r-c),A5:B7)(D6,D7)";
            sheet.Cells["D6"].Value = 7;
            sheet.Cells["D7"].Value = 2;
            sheet.Calculate();
            Assert.AreEqual(5d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void LambdaRecursive1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(a,a + LAMBDA(b,b + a)(2))(2)";
            sheet.Calculate();
            Assert.AreEqual(6d, sheet.Cells["A1"].Value);
        }


        [TestMethod]
        public void LambdaAsName1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            package.Workbook.Names.AddFormula("TestFunc", "_xlfn.LAMBDA(_xlpm.x,_xlfn.CONCAT(\"Testfunc: \",_xlpm.x))");
            sheet.Cells["A1"].Formula = "TestFunc(\"Hej hopp\")";
            sheet.Calculate();
            var v = sheet.Cells["A1"].Value;
            Assert.AreEqual("Testfunc: Hej hopp", v);
        }

        [TestMethod]
        public void LambdaAsName2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            package.Workbook.Names.AddFormula("TestFunc", "LAMBDA(x,CONCAT(\"Testfunc: \",x))");
            sheet.Cells["A1"].Formula = "TestFunc(\"Hej hopp\")";
            sheet.Calculate();
            var v = sheet.Cells["A1"].Value;
            Assert.AreEqual("Testfunc: Hej hopp", v);
        }

        [TestMethod]
        public void LambdaAsName3()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            package.Workbook.Names.AddFormula("TestFunc", "LAMBDA(x,y, x/y)");
            sheet.Cells["A1"].Formula = "TestFunc(12,3)";
            sheet.Calculate();
            var v = sheet.Cells["A1"].Value;
            Assert.AreEqual(4d, v);
        }

        [TestMethod]
        public void LetAndLambdaCombined()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LET(x,LAMBDA(x,y,x+1)(1,2),x+1)";
            sheet.Calculate();
            Assert.AreEqual(3d, sheet.Cells["A1"].Value);
        }


        [TestMethod]
        public void LambdaAddressTest1()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Formula = "LAMBDA(a,a)(A1):A3";
            sheet.Calculate();
            var a1 = sheet.Cells["A1"].Value;
            var a2 = sheet.Cells["A2"].Value;
            var a3 = sheet.Cells["A3"].Value;
            Assert.AreEqual(1, a1);
            Assert.AreEqual(2, a2);
            Assert.AreEqual(3, a3);
        }


        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void LambdaCircularReferenceTest()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Formula = "B1";
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Formula = "LAMBDA(a,a)(A1):A3";
            sheet.Calculate();
        }
    }
}
