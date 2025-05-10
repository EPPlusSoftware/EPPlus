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
    [TestClass]
    public class LambdaFunctionTests : TestBase
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


        //Curiously, no issues in release, Crashes in debug.
        [TestMethod]
        public void ArrayAnchorReduce()
        {
            using (var p = OpenPackage("Reduce_ArrayAnchor.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("ReduceSheet");

                ws.Cells["A1"].Formula = "SEQUENCE(2)";
                ws.Cells["C1"].Formula = "LAMBDA(x,REDUCE(\"\",x,LAMBDA(a,v,SORT(x,,1,TRUE))))(ANCHORARRAY(A1))";

                ws.Calculate();

                SaveAndCleanup(p);
            }
        }

        //We have issues reading and then saving ANCHORARRAY or # operator formulas
        //The "Resaved" file does not output an array but only the first value
        //Also see next test
        [TestMethod]
        public void ArrayAnchorReduce_Resave()
        {
            using (var p = OpenPackage("Reduce_ArrayAnchor_Original.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("ReduceSheet");

                ws.Cells["A1"].Formula = "SEQUENCE(2)";
                ws.Cells["C1"].Formula = "LAMBDA(x,REDUCE(\"\",x,LAMBDA(a,v,SORT(x,,1,TRUE))))(ANCHORARRAY(A1))";

                ws.Calculate();

                SaveAndCleanup(p);
            }

            using (var p = OpenPackage("Reduce_ArrayAnchor_Original.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.ClearFormulaValues();

                ws.Calculate();

                var newName = GetOutputFile("", "Reduce_ArrayAnchor_Resaved.xlsx").FullName;
                p.SaveAs(newName);
            }
        }

        [TestMethod]
        public void LambdaRangeArg_Simple()
        {
            using var p = OpenPackage("LambdaFillDown_Generated_Simple.xlsx", true);
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["C3"].Value = 7d;
            sheet.Cells["C4"].Value = 8d;
            sheet.Cells["D3"].Value = 7d;
            sheet.Cells["D4"].Value = 9d;

            //sheet.Cells["G3"].Formula = "LAMBDA(range, SCAN(\"\", range, LAMBDA(a,v, IF(v = \"\", a, v))))(C3:D4)";
            sheet.Cells["G3"].Formula = "LAMBDA(range,range)(C3:D4)";
            sheet.Calculate();
            Assert.AreEqual(7d, sheet.Cells["G3"].Value, "G3 was not 6 as expected");
            Assert.AreEqual(8d, sheet.Cells["G4"].Value, "G4 was not 7 as expected");
            Assert.AreEqual(7d, sheet.Cells["H3"].Value);
            Assert.AreEqual(9d, sheet.Cells["H4"].Value);

            SaveAndCleanup(p);
            p.Dispose();

            using var p2 = OpenPackage("LambdaFillDown_Generated_Simple.xlsx", false);
            var ws = p2.Workbook.Worksheets[0];
            ws.ClearFormulaValues();

            ws.Calculate();

            var newName = GetOutputFile("", "LambdaFillDown_Generated_Simple_Resaved.xlsx").FullName;
            p2.SaveAs(newName);

        }

        //All array input/output looks to have issues when reading a file.
        //At least if double lambda
        [TestMethod]
        public void MakingLambda_FillDown_RESAVE()
        {
            using (var p = OpenPackage("LambdaFillDown_Generated.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = p.Workbook.Worksheets.Add("FillDown");

                ws.Cells["C3:D4"].Formula = "ROW()+COLUMN()";
                ws.Cells["G3"].Formula = "LAMBDA(range, SCAN(\"\", range, LAMBDA(a,v, IF(v = \"\", a, v))))(C3:D4)";

                ws.Calculate();
                //Output file looks like expected.
                Assert.AreEqual(6d, ws.Cells["G3"].Value, "G3 was not 6 as expected");
                Assert.AreEqual(7d, ws.Cells["G4"].Value, "G4 was not 7 as expected");
                Assert.AreEqual(7d, ws.Cells["H3"].Value, "H3 was not 7 as expected");
                Assert.AreEqual(8d, ws.Cells["H4"].Value, "H4 was not 8 as ");
                SaveAndCleanup(p);
            }

            using (var p = OpenPackage("LambdaFillDown_Generated.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.ClearFormulaValues();

                ws.Calculate();

                //Re-read file does not.
                var newName = GetOutputFile("", "LambdaFillDown_Generated_Resaved.xlsx").FullName;
                p.SaveAs(newName);
            }
        }

        [TestMethod]
        public void TokenizeArray()
        {
            //var f = "xlfn.LAMBDA(_xlpm.Text_to_Change,_xlpm.Substitution_Table, _xlfn.LET( _xlpm.A, \" \"&_xlpm.Text_to_Change&\" \", _xlpm.B, TRIM(_xlpm.Substitution_Table), _xlpm.Prefix, {\"-\",\"\"\"\",\"'\",\" \"}, _xlpm.Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"}, _xlpm.Frm_1, _xlfn.TOCOL(_xlpm.Prefix & _xlfn.TOCOL(_xlfn.CHOOSECOLS(_xlpm.B, 1) & _xlpm.Suffix)), _xlpm.Frm_2, _xlfn.VSTACK(UPPER(_xlpm.Frm_1), LOWER(_xlpm.Frm_1), PROPER(_xlpm.Frm_1)), _xlpm.To_1, _xlfn.TOCOL(_xlpm.Prefix & _xlfn.TOCOL(_xlfn.CHOOSECOLS(_xlpm.B, 2) & _xlpm.Suffix)), _xlpm.To_2, _xlfn.VSTACK(UPPER(_xlpm.To_1), LOWER(_xlpm.To_1), PROPER(_xlpm.To_1)), _xlpm.Output, _xlfn.REDUCE(_xlpm.A, _xlfn.SEQUENCE(ROWS(_xlpm.To_2)), _xlfn.LAMBDA(_xlpm.X,_xlpm.Y, SUBSTITUTE(_xlpm.X, INDEX(_xlpm.Frm_2, _xlpm.Y), INDEX(_xlpm.To_2, _xlpm.Y)))), TRIM(_xlpm.Output)))(E11,J12:K16)";
            //var f2 = "LET(x, {\"-\",\"\"\"\",\"'\",\" \"}, 1)";
            //var tokens = SourceCodeTokenizer.Default.Tokenize(f2);
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LET(x, {\"-\"; \"a\" }, \"item: \" & x)";
            sheet.Calculate();
            Assert.AreEqual("item: -", sheet.Cells["A1"].Value);
            Assert.AreEqual("item: a", sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void TokenizeArray2()
        {
            //var f = "xlfn.LAMBDA(_xlpm.Text_to_Change,_xlpm.Substitution_Table, _xlfn.LET( _xlpm.A, \" \"&_xlpm.Text_to_Change&\" \", _xlpm.B, TRIM(_xlpm.Substitution_Table), _xlpm.Prefix, {\"-\",\"\"\"\",\"'\",\" \"}, _xlpm.Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"}, _xlpm.Frm_1, _xlfn.TOCOL(_xlpm.Prefix & _xlfn.TOCOL(_xlfn.CHOOSECOLS(_xlpm.B, 1) & _xlpm.Suffix)), _xlpm.Frm_2, _xlfn.VSTACK(UPPER(_xlpm.Frm_1), LOWER(_xlpm.Frm_1), PROPER(_xlpm.Frm_1)), _xlpm.To_1, _xlfn.TOCOL(_xlpm.Prefix & _xlfn.TOCOL(_xlfn.CHOOSECOLS(_xlpm.B, 2) & _xlpm.Suffix)), _xlpm.To_2, _xlfn.VSTACK(UPPER(_xlpm.To_1), LOWER(_xlpm.To_1), PROPER(_xlpm.To_1)), _xlpm.Output, _xlfn.REDUCE(_xlpm.A, _xlfn.SEQUENCE(ROWS(_xlpm.To_2)), _xlfn.LAMBDA(_xlpm.X,_xlpm.Y, SUBSTITUTE(_xlpm.X, INDEX(_xlpm.Frm_2, _xlpm.Y), INDEX(_xlpm.To_2, _xlpm.Y)))), TRIM(_xlpm.Output)))(E11,J12:K16)";
            //var f2 = "LET(x, {\"-\",\"\"\"\",\"'\",\" \"}, 1)";
            //var tokens = SourceCodeTokenizer.Default.Tokenize(f2);
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["C11"].Value = "Input";
            sheet.Cells["E11"].Value = "A lizard mage, deals damage by using it's image. Such Imagination.";

            sheet.Cells["J11"].Value = "FROM";
            sheet.Cells["K11"].Value = "TO";

            var fromValues = new List<string>(["mage", "image", "damage", "imagination", "Lizard"]);
            var toValues = new List<string>(["wizard", "iwizard", "dawizard", "iwizardnation", "Gizzard"]);

            sheet.Cells["J12:J16"].LoadFromCollection(fromValues);
            sheet.Cells["K12:K16"].LoadFromCollection(toValues);
            sheet.Cells["C13"].Value = "Output";
            //sheet.Cells["A1"].Formula = "LAMBDA(Text_to_Change, Substitution_table, LET(x, Text_to_Change, x))(E11, J12:J16)";
            //sheet.Cells["A1"].Formula = "LAMBDA(Text_to_Change, Substitution_table, LET(A, Text_to_Change, B, TRIM(Substitution_table), A))(E11, J12:K16)";
            //sheet.Cells["O22"].Formula = "LAMBDA(Text_to_Change, Substitution_table, LET(A, \" \" & Text_to_Change & \" \", B, TRIM(Substitution_table),Prefix,{\"-\";\"\"\"\";\"'\";\" \"}, Suffix, {\"-\";\"\"\"\";\"'\";\" \";\".\";\",\";\":\";\";\";\"=\";\"?\";\"!\"},Frm_1, TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)), Frm_1))(E11, J12:K16)";
            sheet.Cells["O22"].Formula = GetLetFormula7();
            //var tkns = SourceCodeTokenizer.Default.Tokenize(sheet.Cells["O22"].Formula);

            //sheet.Cells["V3"].Formula = "{\"-\",\"\"\"\",\"'\",\" \"} & _xlfn.TOCOL(_xlfn.CHOOSECOLS(J12:K16, 1))";


            sheet.Cells["E11"].Value = "A lizard mage, deals damage by using it's image. Such Imagination.";
            sheet.Calculate();

            // GetLetFormula3
            //Assert.AreEqual("-mage-", sheet.Cells["O22"].Value);
            //Assert.AreEqual("\"mage-", sheet.Cells["O23"].Value);
            //Assert.AreEqual("'mage-", sheet.Cells["O24"].Value);
            //Assert.AreEqual("-image-", sheet.Cells["O66"].Value);
            //Assert.AreEqual(" Lizard!", sheet.Cells["O241"].Value);
            //Assert.IsNull(sheet.Cells["O242"].Value);

            // GetLetFormula4
            //Assert.AreEqual("-MAGE-", sheet.Cells["O22"].Value);
            //Assert.AreEqual(" Lizard!", sheet.Cells["O681"].Value);
            //Assert.IsNull(sheet.Cells["O682"].Value);

            // GetLetFormula5
            //Assert.AreEqual("-wizard-", sheet.Cells["O22"].Value);
            //Assert.AreEqual(" Gizzard!", sheet.Cells["O241"].Value);
            //Assert.IsNull(sheet.Cells["O242"].Value);

            // GetLetFormula6
            //Assert.AreEqual("-WIZARD-", sheet.Cells["O22"].Value);
            //Assert.AreEqual(" Gizzard!", sheet.Cells["O681"].Value);
            //Assert.IsNull(sheet.Cells["O682"].Value);

            Assert.AreEqual("A gizzard wizard, deals dawizard by using it's iwizard. Such Iwizardnation.", sheet.Cells["O22"].Value);
        }

        private string GetLetFormula1()
        {
            var sb = new StringBuilder();
            sb.Append("LAMBDA(Text_to_Change,Substitution_table,");
            sb.Append("LET(A, \" \" & Text_to_Change & \" \",");
            sb.Append("B, TRIM(Substitution_table),");
            sb.Append("Prefix, {\"-\",\"\"\"\",\"'\",\" \"},");
            sb.Append("Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"},");
            sb.Append("Frm_1, TOCOL(CHOOSECOLS(B, 1)),");
            sb.Append("Frm_1))(E11, J12:K16)");
            return sb.ToString();
        }

        private string GetLetFormula2()
        {
            var sb = new StringBuilder();
            sb.Append("LAMBDA(Text_to_Change,Substitution_table,");
            sb.Append("LET(A, \" \" & Text_to_Change & \" \",");
            sb.Append("B, TRIM(Substitution_table),");
            sb.Append("Prefix, {\"-\",\"\"\"\",\"'\",\" \"},");
            sb.Append("Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"},");
            sb.Append("Frm_1, TOCOL(CHOOSECOLS(B, 1) & Suffix),");
            sb.Append("Frm_1))(E11, J12:K16)");
            return sb.ToString();
        }


        private string GetLetFormula3()
        {
            var sb = new StringBuilder();
            sb.Append("LAMBDA(Text_to_Change,Substitution_table,");
            sb.Append("LET(A, \" \" & Text_to_Change & \" \",");
            sb.Append("B, TRIM(Substitution_table),");
            sb.Append("Prefix, {\"-\",\"\"\"\",\"'\",\" \"},");
            sb.Append("Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"},");
            sb.Append("Frm_1,  TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)),");
            sb.Append("Frm_1))(E11, J12:K16)");
            return sb.ToString();
        }

        private string GetLetFormula4()
        {
            var sb = new StringBuilder();
            sb.Append("LAMBDA(Text_to_Change,Substitution_table,");
            sb.Append("LET(A, \" \" & Text_to_Change & \" \",");
            sb.Append("B, TRIM(Substitution_table),");
            sb.Append("Prefix, {\"-\",\"\"\"\",\"'\",\" \"},");
            sb.Append("Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"},");
            sb.Append("Frm_1,  TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)),");
            sb.Append("Frm_2,  VSTACK(UPPER(Frm_1), LOWER(Frm_1), PROPER(Frm_1)),");
            sb.Append("Frm_2))(E11, J12:K16)");
            return sb.ToString();
        }

        private string GetLetFormula5()
        {
            var sb = new StringBuilder();
            sb.Append("LAMBDA(Text_to_Change,Substitution_table,");
            sb.Append("LET(A, \" \" & Text_to_Change & \" \",");
            sb.Append("B, TRIM(Substitution_table),");
            sb.Append("Prefix, {\"-\",\"\"\"\",\"'\",\" \"},");
            sb.Append("Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"},");
            sb.Append("Frm_1,  TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)),");
            sb.Append("Frm_2,  VSTACK(UPPER(Frm_1), LOWER(Frm_1), PROPER(Frm_1)),");
            sb.Append("To_1,   TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 2) & Suffix)),");
            sb.Append("To_1))(E11, J12:K16)");
            return sb.ToString();
        }

        private string GetLetFormula6()
        {
            var sb = new StringBuilder();
            sb.Append("LAMBDA(Text_to_Change,Substitution_table,");
            sb.Append("LET(A, \" \" & Text_to_Change & \" \",");
            sb.Append("B, TRIM(Substitution_table),");
            sb.Append("Prefix, {\"-\",\"\"\"\",\"'\",\" \"},");
            sb.Append("Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"},");
            sb.Append("Frm_1,  TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)),");
            sb.Append("Frm_2,  VSTACK(UPPER(Frm_1), LOWER(Frm_1), PROPER(Frm_1)),");
            sb.Append("To_1,   TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 2) & Suffix)),");
            sb.Append("To_2,   VSTACK(UPPER(To_1), LOWER(To_1), PROPER(To_1)),");
            sb.Append("To_2))(E11, J12:K16)");
            return sb.ToString();
        }


        private string GetLetFormula7()
        {
            var sb = new StringBuilder();
            sb.Append("LAMBDA(Text_to_Change,Substitution_table,");
            sb.Append("LET(A, \" \" & Text_to_Change & \" \",");
            sb.Append("B, TRIM(Substitution_table),");
            sb.Append("Prefix, {\"-\",\"\"\"\",\"'\",\" \"},");
            sb.Append("Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"},");
            sb.Append("Frm_1,  TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)),");
            sb.Append("Frm_2,  VSTACK(UPPER(Frm_1), LOWER(Frm_1), PROPER(Frm_1)),");
            sb.Append("To_1,   TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 2) & Suffix)),");
            sb.Append("To_2,   VSTACK(UPPER(To_1), LOWER(To_1), PROPER(To_1)),");
            sb.Append("Output, REDUCE(A, SEQUENCE(ROWS(To_2)), LAMBDA(X,Y,");
            sb.Append("SUBSTITUTE(X, INDEX(Frm_2, Y), INDEX(To_2, Y)))),");
            sb.Append("TRIM(Output)))(E11, J12:K16)");
            return sb.ToString();
        }

        [TestMethod]
        public void ShouldTokenizeParameters()
        {
            //var f = "LAMBDA(Text_to_Change, Substitution_table, LET(A, \" \" & Text_to_Change & \" \", B, TRIM(Substitution_table),Prefix,{\"-\";\"\"\"\";\"'\";\" \"}, Suffix, {\"-\";\"\"\"\";\"'\";\" \";\".\";\",\";\":\";\";\";\"=\";\"?\";\"!\"},Frm_1, TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)), Frm_1))(E11, J12:K16)";
            var f = "LET(B, 1, Suffix, 2, Frm_1, TOCOL(TOCOL(CHOOSECOLS(B, 1) & Suffix)), Frm_1)";
            var tokens = SourceCodeTokenizer.Default.Tokenize(f);
            Assert.AreEqual(TokenType.ParameterVariable, tokens[23].TokenType);
        }


        [TestMethod]
        public void DaWizard_Generated()
        {
            using (var p = OpenPackage("DaWizard_Generated.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("MageSheet");

                ws.Cells["C11"].Value = "Input";
                ws.Cells["E11"].Value = "A lizard mage, deals damage by using it's image. Such Imagination.";

                ws.Cells["J11"].Value = "FROM";
                ws.Cells["K11"].Value = "TO";

                var fromValues = new List<string>(["mage", "image", "damage", "imagination", "Lizard"]);
                var toValues = new List<string>(["wizard", "iwizard", "dawizard", "iwizardnation", "Gizzard"]);

                ws.Cells["J12:J16"].LoadFromCollection(fromValues);
                ws.Cells["K12:K16"].LoadFromCollection(toValues);
                ws.Cells["C13"].Value = "Output";

                //ws.Cells["C3"].Formula = "LAMBDA(OriginalText,WordSwapTable, LET( A,\"  \"&OriginalText&\"  \", B, TRIM(WordSwapTable), Prefix , {\"-\",\"\"\"\",\"'\",\" \"}, Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"}, Frm_1,  TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)), Frm_2,  VSTACK(UPPER(Frm_1), LOWER(Frm_1), PROPER(Frm_1)), To_1, TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 2) & Suffix)), To_2, VSTACK(UPPER(To_1), LOWER(To_1), PROPER(To_1)), Output, REDUCE(A, SEQUENCE(ROWS(To_2)), LAMBDA(X,Y, SUBSTITUTE(X, INDEX(Frm_2, Y), INDEX(To_2, Y)))), TRIM(Output)))(C1,F3:G7)";
                ws.Cells["E13"].Formula = "LAMBDA(Text_to_Change,Substitution_Table, LET(A, \" \"&Text_to_Change&\" \", B, TRIM(Substitution_Table), Prefix, {\"-\",\"\"\"\",\"'\",\" \"}, Suffix, {\"-\",\"\"\"\",\"'\",\" \",\".\",\",\",\":\",\";\",\"=\",\"?\",\"!\"}, Frm_1, TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 1) & Suffix)), Frm_2, VSTACK(UPPER(Frm_1), LOWER(Frm_1), PROPER(Frm_1)), To_1, TOCOL(Prefix & TOCOL(CHOOSECOLS(B, 2) & Suffix)), To_2, VSTACK(UPPER(To_1), LOWER(To_1), PROPER(To_1)), Output, REDUCE(A, SEQUENCE(ROWS(To_2)), LAMBDA(X,Y, SUBSTITUTE(X, INDEX(Frm_2, Y), INDEX(To_2, Y)))), TRIM(Output)))(E11,J12:K16)";
                ws.Calculate();

                var val = ws.GetValueInner(ws.Cells["E13"].Start.Row, ws.Cells["E13"].Start.Column);

                //Assert.AreEqual("A gizzard wizard, deals dawizard by using it's iwizard. Such Iwizardnation.", val);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ALotOfParametersForLamda()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Shieet 1");

            for(int i =1;i<=254;i++)
            {
                ws.Cells["A" + i].Value = i;
            }

            //=LAMBDA(argu1;argu2;argu3;argu4;argu5;argu6;argu7;argu8;argu9;argu10;argu11;argu12;argu13;argu14;argu15;argu16;argu17;argu18;argu19;argu20;argu21;argu22;argu23;argu24;argu25;argu26;argu27;argu28;argu29;argu30;argu31;argu32;argu33;argu34;argu35;argu36;argu37;argu38;argu39;argu40;argu41;argu42;argu43;argu44;argu45;argu46;argu47;argu48;argu49;argu50;argu51;argu52;argu53;argu54;argu55;argu56;argu57;argu58;argu59;argu60;argu61;argu62;argu63;argu64;argu65;argu66;argu67;argu68;argu69;argu70;argu71;argu72;argu73;argu74;argu75;argu76;argu77;argu78;argu79;argu80;argu81;argu82;argu83;argu84;argu85;argu86;argu87;argu88;argu89;argu90;argu91;argu92;argu93;argu94;argu95;argu96;argu97;argu98;argu99;argu100;argu101;argu102;argu103;argu104;argu105;argu106;argu107;argu108;argu109;argu110;argu111;argu112;argu113;argu114;argu115;argu116;argu117;argu118;argu119;argu120;argu121;argu122;argu123;argu124;argu125;argu126;argu127;argu128;argu129;argu130;argu131;argu132;argu133;argu134;argu135;argu136;argu137;argu138;argu139;argu140;argu141;argu142;argu143;argu144;argu145;argu146;argu147;argu148;argu149;argu150;argu151;argu152;argu153;argu154;argu155;argu156;argu157;argu158;argu159;argu160;argu161;argu162;argu163;argu164;argu165;argu166;argu167;argu168;argu169;argu170;argu171;argu172;argu173;argu174;argu175;argu176;argu177;argu178;argu179;argu180;argu181;argu182;argu183;argu184;argu185;argu186;argu187;argu188;argu189;argu190;argu191;argu192;argu193;argu194;argu195;argu196;argu197;argu198;argu199;argu200;argu201;argu202;argu203;argu204;argu205;argu206;argu207;argu208;argu209;argu210;argu211;argu212;argu213;argu214;argu215;argu216;argu217;argu218;argu219;argu220;argu221;argu222;argu223;argu224;argu225;argu226;argu227;argu228;argu229;argu230;argu231;argu232;argu233;argu234;argu235;argu236;argu237;argu238;argu239;argu240;argu241;argu242;argu243;argu244;argu245;argu246;argu247;argu248;argu249;argu250;argu251;argu252;argu253;argu1+argu2+argu3+argu4+argu5+argu6+argu7+argu8+argu9+argu10+argu11+argu12+argu13+argu14+argu15+argu16+argu17+argu18+argu19+argu20+argu21+argu22+argu23+argu24+argu25+argu26+argu27+argu28+argu29+argu30+argu31+argu32+argu33+argu34+argu35+argu36+argu37+argu38+argu39+argu40+argu41+argu42+argu43+argu44+argu45+argu46+argu47+argu48+argu49+argu50+argu51+argu52+argu53+argu54+argu55+argu56+argu57+argu58+argu59+argu60+argu61+argu62+argu63+argu64+argu65+argu66+argu67+argu68+argu69+argu70+argu71+argu72+argu73+argu74+argu75+argu76+argu77+argu78+argu79+argu80+argu81+argu82+argu83+argu84+argu85+argu86+argu87+argu88+argu89+argu90+argu91+argu92+argu93+argu94+argu95+argu96+argu97+argu98+argu99+argu100+argu101+argu102+argu103+argu104+argu105+argu106+argu107+argu108+argu109+argu110+argu111+argu112+argu113+argu114+argu115+argu116+argu117+argu118+argu119+argu120+argu121+argu122+argu123+argu124+argu125+argu126+argu127+argu128+argu129+argu130+argu131+argu132+argu133+argu134+argu135+argu136+argu137+argu138+argu139+argu140+argu141+argu142+argu143+argu144+argu145+argu146+argu147+argu148+argu149+argu150+argu151+argu152+argu153+argu154+argu155+argu156+argu157+argu158+argu159+argu160+argu161+argu162+argu163+argu164+argu165+argu166+argu167+argu168+argu169+argu170+argu171+argu172+argu173+argu174+argu175+argu176+argu177+argu178+argu179+argu180+argu181+argu182+argu183+argu184+argu185+argu186+argu187+argu188+argu189+argu190+argu191+argu192+argu193+argu194+argu195+argu196+argu197+argu198+argu199+argu200+argu201+argu202+argu203+argu204+argu205+argu206+argu207+argu208+argu209+argu210+argu211+argu212+argu213+argu214+argu215+argu216+argu217+argu218+argu219+argu220+argu221+argu222+argu223+argu224+argu225+argu226+argu227+argu228+argu229+argu230+argu231+argu232+argu233+argu234+argu235+argu236+argu237+argu238+argu239+argu240+argu241+argu242+argu243+argu244+argu245+argu246+argu247+argu248+argu249+argu250+argu251+argu252+argu253)(A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12;A13;A14;A15;A16;A17;A18;A19;A20;A21;A22;A23;A24;A25;A26;A27;A28;A29;A30;A31;A32;A33;A34;A35;A36;A37;A38;A39;A40;A41;A42;A43;A44;A45;A46;A47;A48;A49;A50;A51;A52;A53;A54;A55;A56;A57;A58;A59;A60;A61;A62;A63;A64;A65;A66;A67;A68;A69;A70;A71;A72;A73;A74;A75;A76;A77;A78;A79;A80;A81;A82;A83;A84;A85;A86;A87;A88;A89;A90;A91;A92;A93;A94;A95;A96;A97;A98;A99;A100;A101;A102;A103;A104;A105;A106;A107;A108;A109;A110;A111;A112;A113;A114;A115;A116;A117;A118;A119;A120;A121;A122;A123;A124;A125;A126;A127;A128;A129;A130;A131;A132;A133;A134;A135;A136;A137;A138; A139;A140;A141;A142;A143;A144;A145;A146;A147;A148;A149;A150;A151;A152;A153;A154;A155;A156;A157;A158;A159;A160;A161;A162;A163;A164;A165;A166;A167;A168;A169;A170;A171;A172;A173;A174;A175;A176;A177;A178;A179;A180;A181;A182;A183;A184;A185;A186;A187;A188;A189;A190;A191;A192;A193;A194;A195;A196;A197;A198;A199;A200;A201;A202;A203;A204;A205;A206;A207;A208;A209;A210;A211;A212;A213;A214;A215;A216;A217;A218;A219;A220;A211;A222;A223;A224;A225;A226;A227;A228;A229;A230;A231;A232;A233;A234;A235;A236;A237;A238;A239;A240;A241;A242;A243;A244;A245;A246;A247;A248;A249;A250;A251;A252;A253)

            ws.Cells["C1"].Formula = "=LAMBDA(argu1,argu2,argu3,argu4,argu5,argu6,argu7,argu8,argu9,argu10,argu11,argu12,argu13,argu14,argu15,argu16,argu17,argu18,argu19,argu20,argu21,argu22,argu23,argu24,argu25,argu26,argu27,argu28,argu29,argu30,argu31,argu32,argu33,argu34,argu35,argu36,argu37,argu38,argu39,argu40,argu41,argu42,argu43,argu44,argu45,argu46,argu47,argu48,argu49,argu50,argu51,argu52,argu53,argu54,argu55,argu56,argu57,argu58,argu59,argu60,argu61,argu62,argu63,argu64,argu65,argu66,argu67,argu68,argu69,argu70,argu71,argu72,argu73,argu74,argu75,argu76,argu77,argu78,argu79,argu80,argu81,argu82,argu83,argu84,argu85,argu86,argu87,argu88,argu89,argu90,argu91,argu92,argu93,argu94,argu95,argu96,argu97,argu98,argu99,argu100,argu101,argu102,argu103,argu104,argu105,argu106,argu107,argu108,argu109,argu110,argu111,argu112,argu113,argu114,argu115,argu116,argu117,argu118,argu119,argu120,argu121,argu122,argu123,argu124,argu125,argu126,argu127,argu128,argu129,argu130,argu131,argu132,argu133,argu134,argu135,argu136,argu137,argu138,argu139,argu140,argu141,argu142,argu143,argu144,argu145,argu146,argu147,argu148,argu149,argu150,argu151,argu152,argu153,argu154,argu155,argu156,argu157,argu158,argu159,argu160,argu161,argu162,argu163,argu164,argu165,argu166,argu167,argu168,argu169,argu170,argu171,argu172,argu173,argu174,argu175,argu176,argu177,argu178,argu179,argu180,argu181,argu182,argu183,argu184,argu185,argu186,argu187,argu188,argu189,argu190,argu191,argu192,argu193,argu194,argu195,argu196,argu197,argu198,argu199,argu200,argu201,argu202,argu203,argu204,argu205,argu206,argu207,argu208,argu209,argu210,argu211,argu212,argu213,argu214,argu215,argu216,argu217,argu218,argu219,argu220,argu221,argu222,argu223,argu224,argu225,argu226,argu227,argu228,argu229,argu230,argu231,argu232,argu233,argu234,argu235,argu236,argu237,argu238,argu239,argu240,argu241,argu242,argu243,argu244,argu245,argu246,argu247,argu248,argu249,argu250,argu251,argu252,argu253,argu254," +
                                             "argu1+argu2+argu3+argu4+argu5+argu6+argu7+argu8+argu9+argu10+argu11+argu12+argu13+argu14+argu15+argu16+argu17+argu18+argu19+argu20+argu21+argu22+argu23+argu24+argu25+argu26+argu27+argu28+argu29+argu30+argu31+argu32+argu33+argu34+argu35+argu36+argu37+argu38+argu39+argu40+argu41+argu42+argu43+argu44+argu45+argu46+argu47+argu48+argu49+argu50+argu51+argu52+argu53+argu54+argu55+argu56+argu57+argu58+argu59+argu60+argu61+argu62+argu63+argu64+argu65+argu66+argu67+argu68+argu69+argu70+argu71+argu72+argu73+argu74+argu75+argu76+argu77+argu78+argu79+argu80+argu81+argu82+argu83+argu84+argu85+argu86+argu87+argu88+argu89+argu90+argu91+argu92+argu93+argu94+argu95+argu96+argu97+argu98+argu99+argu100+argu101+argu102+argu103+argu104+argu105+argu106+argu107+argu108+argu109+argu110+argu111+argu112+argu113+argu114+argu115+argu116+argu117+argu118+argu119+argu120+argu121+argu122+argu123+argu124+argu125+argu126+argu127+argu128+argu129+argu130+argu131+argu132+argu133+argu134+argu135+argu136+argu137+argu138+argu139+argu140+argu141+argu142+argu143+argu144+argu145+argu146+argu147+argu148+argu149+argu150+argu151+argu152+argu153+argu154+argu155+argu156+argu157+argu158+argu159+argu160+argu161+argu162+argu163+argu164+argu165+argu166+argu167+argu168+argu169+argu170+argu171+argu172+argu173+argu174+argu175+argu176+argu177+argu178+argu179+argu180+argu181+argu182+argu183+argu184+argu185+argu186+argu187+argu188+argu189+argu190+argu191+argu192+argu193+argu194+argu195+argu196+argu197+argu198+argu199+argu200+argu201+argu202+argu203+argu204+argu205+argu206+argu207+argu208+argu209+argu210+argu211+argu212+argu213+argu214+argu215+argu216+argu217+argu218+argu219+argu220+argu221+argu222+argu223+argu224+argu225+argu226+argu227+argu228+argu229+argu230+argu231+argu232+argu233+argu234+argu235+argu236+argu237+argu238+argu239+argu240+argu241+argu242+argu243+argu244+argu245+argu246+argu247+argu248+argu249+argu250+argu251+argu252+argu253+254)" +
                                             "(A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,A16,A17,A18,A19,A20,A21,A22,A23,A24,A25,A26,A27,A28,A29,A30,A31,A32,A33,A34,A35,A36,A37,A38,A39,A40,A41,A42,A43,A44,A45,A46,A47,A48,A49,A50,A51,A52,A53,A54,A55,A56,A57,A58,A59,A60,A61,A62,A63,A64,A65,A66,A67,A68,A69,A70,A71,A72,A73,A74,A75,A76,A77,A78,A79,A80,A81,A82,A83,A84,A85,A86,A87,A88,A89,A90,A91,A92,A93,A94,A95,A96,A97,A98,A99,A100,A101,A102,A103,A104,A105,A106,A107,A108,A109,A110,A111,A112,A113,A114,A115,A116,A117,A118,A119,A120,A121,A122,A123,A124,A125,A126,A127,A128,A129,A130,A131,A132,A133,A134,A135,A136,A137,A138,A139,A140,A141,A142,A143,A144,A145,A146,A147,A148,A149,A150,A151,A152,A153,A154,A155,A156,A157,A158,A159,A160,A161,A162,A163,A164,A165,A166,A167,A168,A169,A170,A171,A172,A173,A174,A175,A176,A177,A178,A179,A180,A181,A182,A183,A184,A185,A186,A187,A188,A189,A190,A191,A192,A193,A194,A195,A196,A197,A198,A199,A200,A201,A202,A203,A204,A205,A206,A207,A208,A209,A210,A211,A212,A213,A214,A215,A216,A217,A218,A219,A220,A211,A222,A223,A224,A225,A226,A227,A228,A229,A230,A231,A232,A233,A234,A235,A236,A237,A238,A239,A240,A241,A242,A243,A244,A245,A246,A247,A248,A249,A250,A251,A252,A253,A254)";
            ws.Calculate();
            var result = ws.Cells["C1"].Value;
            SaveWorkbook("DumbLambdaSum.xlsx", p);
        }

        //Excel Calculates correctly. Epplus gets Value! error
        //Solve first
        [TestMethod]
        public void RecursiveFormulaSimple()
        {
            using (var p = OpenPackage("RecursiveSimple.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                //Add own Lambda Factorial function via Let and IF
                ws.Names.AddFormula("Factorial2", "LAMBDA(input, LET(n, input, IF(n = 0, 1, n * Sheet1!Factorial2(n - 1))))");

                ws.Cells["A1"].Formula = "Factorial2(4)";

                ws.Calculate();

                var epplusValue = ws.Cells["A1"].Value;

                SaveAndCleanup(p);

                Assert.AreEqual(24d, epplusValue);
            }
        }

        //Same as RecursiveFormulaSimple but without "Sheet1!" specifed. Name error in both epplus and excel
        //If Formula is set in Excel it automatically adds "Sheet1!" in workbook.xml. Epplus should also realise this.
        [TestMethod]
        public void RecursiveFormulaSimple_SheetUnspecified()
        {
            using (var p = OpenPackage("RecursiveSimple_Unspecifed.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                //Add own Lambda Factorial function via Let and IF
                ws.Names.AddFormula("Factorial2", "LAMBDA(input, LET(n, input, IF(n = 0, 1, n * Factorial2(n - 1))))");

                ws.Cells["A1"].Formula = "Factorial2(4)";

                ws.Calculate();

                var epplusValue = ws.Cells["A1"].Value;

                SaveAndCleanup(p);

                Assert.AreEqual(24d, epplusValue);
            }
        }
    }
}
