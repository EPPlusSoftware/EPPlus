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
        public void ALotOfParametersForLamda()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Shieet 1");

            var a1 = ws.Cells["A1"].Value = 1;
            var a2 = ws.Cells["A2"].Value = 2;
            var a3 = ws.Cells["A3"].Value = 3;
            var a4 = ws.Cells["A4"].Value = 4;
            var a5 = ws.Cells["A5"].Value = 5;
            var a6 = ws.Cells["A6"].Value = 6;
            var a7 = ws.Cells["A7"].Value = 7;
            var a8 = ws.Cells["A8"].Value = 8;
            var a9 = ws.Cells["A9"].Value = 9;
            var a10 = ws.Cells["A10"].Value = 10;
            var a11 = ws.Cells["A11"].Value = 11;
            var a12 = ws.Cells["A12"].Value = 12;
            var a13 = ws.Cells["A13"].Value = 13;
            var a14 = ws.Cells["A14"].Value = 14;
            var a15 = ws.Cells["A15"].Value = 15;
            var a16 = ws.Cells["A16"].Value = 16;
            var a17 = ws.Cells["A17"].Value = 17;
            var a18 = ws.Cells["A18"].Value = 18;
            var a19 = ws.Cells["A19"].Value = 19;
            var a20 = ws.Cells["A20"].Value = 20;
            var a21 = ws.Cells["A21"].Value = 21;
            var a22 = ws.Cells["A22"].Value = 22;
            var a23 = ws.Cells["A23"].Value = 23;
            var a24 = ws.Cells["A24"].Value = 24;
            var a25 = ws.Cells["A25"].Value = 25;
            var a26 = ws.Cells["A26"].Value = 26;
            var a27 = ws.Cells["A27"].Value = 27;
            var a28 = ws.Cells["A28"].Value = 28;
            var a29 = ws.Cells["A29"].Value = 29;
            var a30 = ws.Cells["A30"].Value = 30;
            var a31 = ws.Cells["A31"].Value = 31;
            var a32 = ws.Cells["A32"].Value = 32;
            var a33 = ws.Cells["A33"].Value = 33;
            var a34 = ws.Cells["A34"].Value = 34;
            var a35 = ws.Cells["A35"].Value = 35;
            var a36 = ws.Cells["A36"].Value = 36;
            var a37 = ws.Cells["A37"].Value = 37;
            var a38 = ws.Cells["A38"].Value = 38;
            var a39 = ws.Cells["A39"].Value = 39;
            var a40 = ws.Cells["A40"].Value = 40;
            var a41 = ws.Cells["A41"].Value = 41;
            var a42 = ws.Cells["A42"].Value = 42;
            var a43 = ws.Cells["A43"].Value = 43;
            var a44 = ws.Cells["A44"].Value = 44;
            var a45 = ws.Cells["A45"].Value = 45;
            var a46 = ws.Cells["A46"].Value = 46;
            var a47 = ws.Cells["A47"].Value = 47;
            var a48 = ws.Cells["A48"].Value = 48;
            var a49 = ws.Cells["A49"].Value = 49;
            var a50 = ws.Cells["A50"].Value = 50;
            var a51 = ws.Cells["A51"].Value = 51;
            var a52 = ws.Cells["A52"].Value = 52;
            var a53 = ws.Cells["A53"].Value = 53;
            var a54 = ws.Cells["A54"].Value = 54;
            var a55 = ws.Cells["A55"].Value = 55;
            var a56 = ws.Cells["A56"].Value = 56;
            var a57 = ws.Cells["A57"].Value = 57;
            var a58 = ws.Cells["A58"].Value = 58;
            var a59 = ws.Cells["A59"].Value = 59;
            var a60 = ws.Cells["A60"].Value = 60;
            var a61 = ws.Cells["A61"].Value = 61;
            var a62 = ws.Cells["A62"].Value = 62;
            var a63 = ws.Cells["A63"].Value = 63;
            var a64 = ws.Cells["A64"].Value = 64;
            var a65 = ws.Cells["A65"].Value = 65;
            var a66 = ws.Cells["A66"].Value = 66;
            var a67 = ws.Cells["A67"].Value = 67;
            var a68 = ws.Cells["A68"].Value = 68;
            var a69 = ws.Cells["A69"].Value = 69;
            var a70 = ws.Cells["A70"].Value = 70;
            var a71 = ws.Cells["A71"].Value = 71;
            var a72 = ws.Cells["A72"].Value = 72;
            var a73 = ws.Cells["A73"].Value = 73;
            var a74 = ws.Cells["A74"].Value = 74;
            var a75 = ws.Cells["A75"].Value = 75;
            var a76 = ws.Cells["A76"].Value = 76;
            var a77 = ws.Cells["A77"].Value = 77;
            var a78 = ws.Cells["A78"].Value = 78;
            var a79 = ws.Cells["A79"].Value = 79;
            var a80 = ws.Cells["A80"].Value = 80;
            var a81 = ws.Cells["A81"].Value = 81;
            var a82 = ws.Cells["A82"].Value = 82;
            var a83 = ws.Cells["A83"].Value = 83;
            var a84 = ws.Cells["A84"].Value = 84;
            var a85 = ws.Cells["A85"].Value = 85;
            var a86 = ws.Cells["A86"].Value = 86;
            var a87 = ws.Cells["A87"].Value = 87;
            var a88 = ws.Cells["A88"].Value = 88;
            var a89 = ws.Cells["A89"].Value = 89;
            var a90 = ws.Cells["A90"].Value = 90;
            var a91 = ws.Cells["A91"].Value = 91;
            var a92 = ws.Cells["A92"].Value = 92;
            var a93 = ws.Cells["A93"].Value = 93;
            var a94 = ws.Cells["A94"].Value = 94;
            var a95 = ws.Cells["A95"].Value = 95;
            var a96 = ws.Cells["A96"].Value = 96;
            var a97 = ws.Cells["A97"].Value = 97;
            var a98 = ws.Cells["A98"].Value = 98;
            var a99 = ws.Cells["A99"].Value = 99;
            var a100 = ws.Cells["A100"].Value = 100;

            var a101 = ws.Cells["A101"].Value = 101;
            var a102 = ws.Cells["A102"].Value = 102;
            var a103 = ws.Cells["A103"].Value = 103;
            var a104 = ws.Cells["A104"].Value = 104;
            var a105 = ws.Cells["A105"].Value = 105;
            var a106 = ws.Cells["A106"].Value = 106;
            var a107 = ws.Cells["A107"].Value = 107;
            var a108 = ws.Cells["A108"].Value = 108;
            var a109 = ws.Cells["A109"].Value = 109;
            var a110 = ws.Cells["A110"].Value = 110;
            var a111 = ws.Cells["A111"].Value = 111;
            var a112 = ws.Cells["A112"].Value = 112;
            var a113 = ws.Cells["A113"].Value = 113;
            var a114 = ws.Cells["A114"].Value = 114;
            var a115 = ws.Cells["A115"].Value = 115;
            var a116 = ws.Cells["A116"].Value = 116;
            var a117 = ws.Cells["A117"].Value = 117;
            var a118 = ws.Cells["A118"].Value = 118;
            var a119 = ws.Cells["A119"].Value = 119;
            var a120 = ws.Cells["A120"].Value = 120;
            var a121 = ws.Cells["A121"].Value = 121;
            var a122 = ws.Cells["A122"].Value = 122;
            var a123 = ws.Cells["A123"].Value = 123;
            var a124 = ws.Cells["A124"].Value = 124;
            var a125 = ws.Cells["A125"].Value = 125;
            var a126 = ws.Cells["A126"].Value = 126;
            var a127 = ws.Cells["A127"].Value = 127;
            var a128 = ws.Cells["A128"].Value = 128;
            var a129 = ws.Cells["A129"].Value = 129;
            var a130 = ws.Cells["A130"].Value = 130;
            var a131 = ws.Cells["A131"].Value = 131;
            var a132 = ws.Cells["A132"].Value = 132;
            var a133 = ws.Cells["A133"].Value = 133;
            var a134 = ws.Cells["A134"].Value = 134;
            var a135 = ws.Cells["A135"].Value = 135;
            var a136 = ws.Cells["A136"].Value = 136;
            var a137 = ws.Cells["A137"].Value = 137;
            var a138 = ws.Cells["A138"].Value = 138;
            var a139 = ws.Cells["A139"].Value = 139;
            var a140 = ws.Cells["A140"].Value = 140;
            var a141 = ws.Cells["A141"].Value = 141;
            var a142 = ws.Cells["A142"].Value = 142;
            var a143 = ws.Cells["A143"].Value = 143;
            var a144 = ws.Cells["A144"].Value = 144;
            var a145 = ws.Cells["A145"].Value = 145;
            var a146 = ws.Cells["A146"].Value = 146;
            var a147 = ws.Cells["A147"].Value = 147;
            var a148 = ws.Cells["A148"].Value = 148;
            var a149 = ws.Cells["A149"].Value = 149;
            var a150 = ws.Cells["A150"].Value = 150;
            var a151 = ws.Cells["A151"].Value = 151;
            var a152 = ws.Cells["A152"].Value = 152;
            var a153 = ws.Cells["A153"].Value = 153;
            var a154 = ws.Cells["A154"].Value = 154;
            var a155 = ws.Cells["A155"].Value = 155;
            var a156 = ws.Cells["A156"].Value = 156;
            var a157 = ws.Cells["A157"].Value = 157;
            var a158 = ws.Cells["A158"].Value = 158;
            var a159 = ws.Cells["A159"].Value = 159;
            var a160 = ws.Cells["A160"].Value = 160;
            var a161 = ws.Cells["A161"].Value = 161;
            var a162 = ws.Cells["A162"].Value = 162;
            var a163 = ws.Cells["A163"].Value = 163;
            var a164 = ws.Cells["A164"].Value = 164;
            var a165 = ws.Cells["A165"].Value = 165;
            var a166 = ws.Cells["A166"].Value = 166;
            var a167 = ws.Cells["A167"].Value = 167;
            var a168 = ws.Cells["A168"].Value = 168;
            var a169 = ws.Cells["A169"].Value = 169;
            var a170 = ws.Cells["A170"].Value = 170;
            var a171 = ws.Cells["A171"].Value = 171;
            var a172 = ws.Cells["A172"].Value = 172;
            var a173 = ws.Cells["A173"].Value = 173;
            var a174 = ws.Cells["A174"].Value = 174;
            var a175 = ws.Cells["A175"].Value = 175;
            var a176 = ws.Cells["A176"].Value = 176;
            var a177 = ws.Cells["A177"].Value = 177;
            var a178 = ws.Cells["A178"].Value = 178;
            var a179 = ws.Cells["A179"].Value = 179;
            var a180 = ws.Cells["A180"].Value = 180;
            var a181 = ws.Cells["A181"].Value = 181;
            var a182 = ws.Cells["A182"].Value = 182;
            var a183 = ws.Cells["A183"].Value = 183;
            var a184 = ws.Cells["A184"].Value = 184;
            var a185 = ws.Cells["A185"].Value = 185;
            var a186 = ws.Cells["A186"].Value = 186;
            var a187 = ws.Cells["A187"].Value = 187;
            var a188 = ws.Cells["A188"].Value = 188;
            var a189 = ws.Cells["A189"].Value = 189;
            var a190 = ws.Cells["A190"].Value = 190;
            var a191 = ws.Cells["A191"].Value = 191;
            var a192 = ws.Cells["A192"].Value = 192;
            var a193 = ws.Cells["A193"].Value = 193;
            var a194 = ws.Cells["A194"].Value = 194;
            var a195 = ws.Cells["A195"].Value = 195;
            var a196 = ws.Cells["A196"].Value = 196;
            var a197 = ws.Cells["A197"].Value = 197;
            var a198 = ws.Cells["A198"].Value = 198;
            var a199 = ws.Cells["A199"].Value = 199;
            var a200 = ws.Cells["A200"].Value = 200;

            var a201 = ws.Cells["A201"].Value = 201;
            var a202 = ws.Cells["A202"].Value = 202;
            var a203 = ws.Cells["A203"].Value = 203;
            var a204 = ws.Cells["A204"].Value = 204;
            var a205 = ws.Cells["A205"].Value = 205;
            var a206 = ws.Cells["A206"].Value = 206;
            var a207 = ws.Cells["A207"].Value = 207;
            var a208 = ws.Cells["A208"].Value = 208;
            var a209 = ws.Cells["A209"].Value = 209;
            var a210 = ws.Cells["A210"].Value = 210;
            var a211 = ws.Cells["A211"].Value = 211;
            var a212 = ws.Cells["A212"].Value = 212;
            var a213 = ws.Cells["A213"].Value = 213;
            var a214 = ws.Cells["A214"].Value = 214;
            var a215 = ws.Cells["A215"].Value = 215;
            var a216 = ws.Cells["A216"].Value = 216;
            var a217 = ws.Cells["A217"].Value = 217;
            var a218 = ws.Cells["A218"].Value = 218;
            var a219 = ws.Cells["A219"].Value = 219;
            var a220 = ws.Cells["A220"].Value = 220;
            var a221 = ws.Cells["A221"].Value = 221;
            var a222 = ws.Cells["A222"].Value = 222;
            var a223 = ws.Cells["A223"].Value = 223;
            var a224 = ws.Cells["A224"].Value = 224;
            var a225 = ws.Cells["A225"].Value = 225;
            var a226 = ws.Cells["A226"].Value = 226;
            var a227 = ws.Cells["A227"].Value = 227;
            var a228 = ws.Cells["A228"].Value = 228;
            var a229 = ws.Cells["A229"].Value = 229;
            var a230 = ws.Cells["A230"].Value = 230;
            var a231 = ws.Cells["A231"].Value = 231;
            var a232 = ws.Cells["A232"].Value = 232;
            var a233 = ws.Cells["A233"].Value = 233;
            var a234 = ws.Cells["A234"].Value = 234;
            var a235 = ws.Cells["A235"].Value = 235;
            var a236 = ws.Cells["A236"].Value = 236;
            var a237 = ws.Cells["A237"].Value = 237;
            var a238 = ws.Cells["A238"].Value = 238;
            var a239 = ws.Cells["A239"].Value = 239;
            var a240 = ws.Cells["A240"].Value = 240;
            var a241 = ws.Cells["A241"].Value = 241;
            var a242 = ws.Cells["A242"].Value = 242;
            var a243 = ws.Cells["A243"].Value = 243;
            var a244 = ws.Cells["A244"].Value = 244;
            var a245 = ws.Cells["A245"].Value = 245;
            var a246 = ws.Cells["A246"].Value = 246;
            var a247 = ws.Cells["A247"].Value = 247;
            var a248 = ws.Cells["A248"].Value = 248;
            var a259 = ws.Cells["A249"].Value = 249;
            var a250 = ws.Cells["A250"].Value = 250;
            var a251 = ws.Cells["A251"].Value = 251;
            var a252 = ws.Cells["A252"].Value = 252;
            var a253 = ws.Cells["A253"].Value = 253;

            var a254 = ws.Cells["A254"].Value = 254;

            var lamda = ws.Cells["C1"].Formula = "=LAMBDA(argu1;argu2;argu3;argu4;argu5;argu6;argu7;argu8;argu9;argu10;argu11;argu12;argu13;argu14;argu15;argu16;argu17;argu18;argu19;argu20;argu21;argu22;argu23;argu24;argu25;argu26;argu27;argu28;argu29;argu30;argu31;argu32;argu33;argu34;argu35;argu36;argu37;argu38;argu39;argu40;argu41;argu42;argu43;argu44;argu45;argu46;argu47;argu48;argu49;argu50;argu51;argu52;argu53;argu54;argu55;argu56;argu57;argu58;argu59;argu60;argu61;argu62;argu63;argu64;argu65;argu66;argu67;argu68;argu69;argu70;argu71;argu72;argu73;argu74;argu75;argu76;argu77;argu78;argu79;argu80;argu81;argu82;argu83;argu84;argu85;argu86;argu87;argu88;argu89;argu90;argu91;argu92;argu93;argu94;argu95;argu96;argu97;argu98;argu99;argu100;argu101;argu102;argu103;argu104;argu105;argu106;argu107;argu108;argu109;argu110;argu111;argu112;argu113;argu114;argu115;argu116;argu117;argu118;argu119;argu120;argu121;argu122;argu123;argu124;argu125;argu126;argu127;argu128;argu129;argu130;argu131;argu132;argu133;argu134;argu135;argu136;argu137;argu138;argu139;argu140;argu141;argu142;argu143;argu144;argu145;argu146;argu147;argu148;argu149;argu150;argu151;argu152;argu153;argu154;argu155;argu156;argu157;argu158;argu159;argu160;argu161;argu162;argu163;argu164;argu165;argu166;argu167;argu168;argu169;argu170;argu171;argu172;argu173;argu174;argu175;argu176;argu177;argu178;argu179;argu180;argu181;argu182;argu183;argu184;argu185;argu186;argu187;argu188;argu189;argu190;argu191;argu192;argu193;argu194;argu195;argu196;argu197;argu198;argu199;argu200;argu201;argu202;argu203;argu204;argu205;argu206;argu207;argu208;argu209;argu210;argu211;argu212;argu213;argu214;argu215;argu216;argu217;argu218;argu219;argu220;argu221;argu222;argu223;argu224;argu225;argu226;argu227;argu228;argu229;argu230;argu231;argu232;argu233;argu234;argu235;argu236;argu237;argu238;argu239;argu240;argu241;argu242;argu243;argu244;argu245;argu246;argu247;argu248;argu249;argu250;argu251;argu252;argu253;argu1+argu2+argu3+argu4+argu5+argu6+argu7+argu8+argu9+argu10+argu11+argu12+argu13+argu14+argu15+argu16+argu17+argu18+argu19+argu20+argu21+argu22+argu23+argu24+argu25+argu26+argu27+argu28+argu29+argu30+argu31+argu32+argu33+argu34+argu35+argu36+argu37+argu38+argu39+argu40+argu41+argu42+argu43+argu44+argu45+argu46+argu47+argu48+argu49+argu50+argu51+argu52+argu53+argu54+argu55+argu56+argu57+argu58+argu59+argu60+argu61+argu62+argu63+argu64+argu65+argu66+argu67+argu68+argu69+argu70+argu71+argu72+argu73+argu74+argu75+argu76+argu77+argu78+argu79+argu80+argu81+argu82+argu83+argu84+argu85+argu86+argu87+argu88+argu89+argu90+argu91+argu92+argu93+argu94+argu95+argu96+argu97+argu98+argu99+argu100+argu101+argu102+argu103+argu104+argu105+argu106+argu107+argu108+argu109+argu110+argu111+argu112+argu113+argu114+argu115+argu116+argu117+argu118+argu119+argu120+argu121+argu122+argu123+argu124+argu125+argu126+argu127+argu128+argu129+argu130+argu131+argu132+argu133+argu134+argu135+argu136+argu137+argu138+argu139+argu140+argu141+argu142+argu143+argu144+argu145+argu146+argu147+argu148+argu149+argu150+argu151+argu152+argu153+argu154+argu155+argu156+argu157+argu158+argu159+argu60+argu161+argu162+argu163+argu164+argu165+argu166+argu167+argu168+argu169+argu170+argu171+argu172+argu173+argu174+argu175+argu176+argu177+argu178+argu179+argu180+argu181+argu182+argu183+argu184+argu185+argu186+argu187+argu188+argu189+argu190+argu191+argu192+argu193+argu194+argu195+argu196+argu197+argu198+argu199+argu200+argu201+argu202+argu203+argu204+argu205+argu206+argu207+argu208+argu209+argu210+argu211+argu212+argu213+argu214+argu215+argu216+argu217+argu218+argu219+argu220+argu221+argu222+argu223+argu224+argu225+argu226+argu227+argu228+argu229+argu230+argu231+argu232+argu233+argu234+argu235+argu236+argu237+argu238+argu239+argu240+argu241+argu242+argu243+argu244+argu245+argu246+argu247+argu248+argu249+argu250+argu251+argu252+argu253)(A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12;A13;A14;A15;A16;A17;A18;A19;A20;A21;A22;A23;A24;A25;A26;A27;A28;A29;A30;A31;A32;A33;A34;A35;A36;A37;A38;A39;A40;A41;A42;A43;A44;A45;A46;A47;A48;A49;A50;A51;A52;A53;A54;A55;A56;A57;A58;A59;A60;A61;A62;A63;A64;A65;A66;A67;A68;A69;A70;A71;A72;A73;A74;A75;A76;A77;A78;A79;A80;A81;A82;A83;A84;A85;A86;A87;A88;A89;A90;A91;A92;A93;A94;A95;A96;A97;A98;A99;A100;A101;A102;A103;A104;A105;A106;A107;A108;A109;A110;A111;A112;A113;A114;A115;A116;A117;A118;A119;A120;A121;A122;A123;A124;A125;A126;A127;A128;A129;A130;A131;A132;A133;A134;A135;A136;A137;A138; A139;A140;A141;A142;A143;A144;A145;A146;A147;A148;A149;A150;A151;A152;A153;A154;A155;A156;A157;A158;A159;A160;A161;A162;A163;A164;A165;A166;A167;A168;A169;A170;A171;A172;A173;A174;A175;A176;A177;A178;A179;A180;A181;A182;A183;A184;A185;A186;A187;A188;A189;A190;A191;A192;A193;A194;A195;A196;A197;A198;A199;A200;A201;A202;A203;A204;A205;A206;A207;A208;A209;A210;A211;A212;A213;A214;A215;A216;A217;A218;A219;A220;A211;A222;A223;A224;A225;A226;A227;A228;A229;A230;A231;A232;A233;A234;A235;A236;A237;A238;A239;A240;A241;A242;A243;A244;A245;A246;A247;A248;A249;A250;A251;A252;A253)";
            ws.Calculate();
            SaveWorkbook("DumbLambdaSum.xlsx", p);
        }
    }
}
