using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class OptimizedSourceCodeTokenizerTests
    {
        private ISourceCodeTokenizer _tokenizer;
        private static int _iterations;
        [TestInitialize]
        public void Setup()
        {
            //_tokenizer = SourceCodeTokenizer.Default;
            _tokenizer = OptimizedSourceCodeTokenizer.Default;
            _iterations = 1000;
        }

        [TestMethod]
        public void TokenizePerformance()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var formula = "VLOOKUP(CONCAT(ORRange30,$H20,$F$17),Ranking!$A$1:$M$3775,MATCH(\"\"\"Value\"\"\",Ranking!$A$1:$M$1,0),0)";
                //var formula = "(-1+-2*3)*12";
                //RunTokenize(tOld, formula);
                RunTokenize(_tokenizer, formula);

                formula = "(( A1 -(- A2 )-( A3 + A4 + A5 ))/( A6 + A7 + A8 - A9 )* A5 )";
                RunTokenize(_tokenizer, formula);

                formula = "SUM(A1:OFFSET(B1;1;3))";
                //RunTokenize(tOld, formula);
                RunTokenize(_tokenizer, formula);
            }
        }

        private static void RunTokenize(OfficeOpenXml.FormulaParsing.LexicalAnalysis.ISourceCodeTokenizer t, string formula)
        {
            var time = DateTime.Now;
            for (int i = 0; i < _iterations; i++)
            {
                var tokens = t.Tokenize(formula, "sheet1");
            }
            var offset = new TimeSpan((DateTime.Now - time).Ticks);
            Debug.WriteLine(offset.TotalMilliseconds);
        }
        [TestMethod]
        public void ShouldCreateTokensForStringCorrectly()
        {
            var input = "\"abc123\"";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(1, tokens.Count());
            Assert.IsTrue(tokens.First().TokenTypeIsSet(TokenType.StringContent));
        }

        [TestMethod]
        public void ShouldTokenizeStringCorrectly()
        {
            var input = "\"ab(c)d\"";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(1, tokens.Count());
        }

        [TestMethod]
        public void ShouldHandleWhitespaceCorrectly()
        {
            var input = @"""          """;
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(1, tokens.Count());
            Assert.IsTrue(tokens.ElementAt(0).TokenTypeIsSet(TokenType.StringContent));
            Assert.AreEqual(10, tokens.ElementAt(0).Value.Length);
        }

        [TestMethod]
        public void ShouldCreateTokensForFunctionCorrectly()
        {
            var input = "Text(2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(4, tokens.Count());
            Assert.IsTrue(tokens.First().TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens.ElementAt(2).TokenTypeIsSet(TokenType.Integer));
            Assert.AreEqual("2", tokens.ElementAt(2).Value);
            Assert.IsTrue(tokens.Last().TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void ShouldHandleMultipleCharOperatorCorrectly()
        {
            var input = "1 <= 2";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count);
            Assert.AreEqual("<=", tokens.ElementAt(1).Value);
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Operator));
        }

        [TestMethod]
        public void ShouldCreateTokensForEnumerableCorrectly()
        {
            var input = "Text({1;2})";
            var tokens = _tokenizer.Tokenize(input).ToArray();

            Assert.AreEqual(8, tokens.Count());
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.OpeningEnumerable));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.ClosingEnumerable));
        }

        [TestMethod]
        public void ShouldCreateTokensWithStringForEnumerableCorrectly()
        {
            var input = "{\"1\",\"2\"}";
            var tokens = _tokenizer.Tokenize(input).ToArray();

            Assert.AreEqual(5, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.OpeningEnumerable));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.StringContent));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.ClosingEnumerable));
        }

        [TestMethod]
        public void ShouldCreateTokensForExcelAddressCorrectly()
        {
            var input = "Text(A1)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.IsTrue(tokens.ElementAt(2).TokenTypeIsSet(TokenType.CellAddress));
        }

        [TestMethod]
        public void ShouldCreateTokenForPercentAfterDecimal()
        {
            var input = "1,23%";
            var tokens = _tokenizer.Tokenize(input);
            Assert.IsTrue(tokens.Last().TokenTypeIsSet(TokenType.Percent));
        }

        [TestMethod]
        public void ShouldIgnoreTwoSubsequentStringIdentifyers()
        {
            var input = "\"hello\"\"world\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(1, tokens.Count());
            Assert.AreEqual("hello\"world", tokens.ElementAt(0).Value);
        }

        [TestMethod]
        public void ShouldIgnoreTwoSubsequentStringIdentifyers2()
        {
            var input = "\"\"\"\"\"\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.IsTrue(tokens.ElementAt(0).TokenTypeIsSet(TokenType.StringContent));
        }

        [TestMethod]
        public void TokenizerShouldIgnoreOperatorInString()
        {
            var input = "\"*\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.IsTrue(tokens.ElementAt(0).TokenTypeIsSet(TokenType.StringContent));
        }

        [TestMethod]
        public void TokenizerShouldHandleWorksheetNameWithMinus()
        {
            var input = "'A-B'!A1";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(3, tokens.Count);
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.CellAddress));
        }
        [TestMethod]
        public void OffsetInAddressTokensFirst()
        {
            var input = "SUM(OFFSET(A3, -1, 0):A1)";
            var tokens = _tokenizer.Tokenize(input).ToList();
            var tokens2 = OptimizedSourceCodeTokenizer.Default.Tokenize(input).ToList();
            for(int i=0;i<tokens.Count();i++)
            {
                Assert.IsTrue(tokens[i].TokenTypeIsSet(tokens2[i].TokenType));
                Assert.AreEqual(tokens[i].Value, tokens2[i].Value);
            }
        }
        [TestMethod]
        public void TestBug9_12_14()
        {
            //(( W60 -(- W63 )-( W29 + W30 + W31 ))/( W23 + W28 + W42 - W51 )* W4 )
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("test");
                for (var x = 1; x <= 10; x++)
                {
                    ws1.Cells[x, 1].Value = x;
                }

                ws1.Cells["A11"].Formula = "(( A1 -(- A2 )-( A3 + A4 + A5 ))/( A6 + A7 + A8 - A9 )* A5 )";
                //ws1.Cells["A11"].Formula = "(-A2 + 1 )";
                ws1.Calculate();
                var result = ws1.Cells["A11"].Value;
                Assert.AreEqual(-3.75, result);
            }
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSign()
        {
            var input = @"+3-3";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(3, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeIdentifiesDoubleNegator()
        {
            var input = @"--3-3";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(5, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeHandlesPositiveNegator()
        {
            var input = @"+-3-3";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(4, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositive()
        {
            var input = @"-+3-3";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(4, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSignFromFirstFunctionArgument()
        {
            var input = @"SUM(+3-3,5)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Count());

            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSignFromSecondFunctionArgument()
        {
            var input = @"SUM(5,+3-3)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingDoubleNegatorFromFirstFunctionArgument()
        {
            var input = @"SUM(--3-3,5)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(10, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[9].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingDoubleNegatorFromSecondFunctionArgument()
        {
            var input = @"SUM(5,--3-3)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(10, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function), "TokenType was not function");
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis), "TokenType was not OpeningParenthesis");
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer), "TokenType was not Integer 2");
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Negator), "TokenType was not negator 4");
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Negator), "TokenType was not negator 5");
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[9].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesPositiveNegatorAsFirstFunctionArgument()
        {
            var input = @"SUM(+-3-3,5)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositiveAsFirstFunctionArgument()
        {
            var input = @"SUM(-+3-3,5)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesPositiveNegatorAsSecondFunctionArgument()
        {
            var input = @"SUM(5,+-3-3)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositiveAsSecondFunctionArgument()
        {
            var input = @"SUM(5,-+3-3)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Count());
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }
        [TestMethod]
        public void TokenizeWorksheetName()
        {
            var input = @"sheetname!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(3, tokens.Count());
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.NameValue));
        }

        [TestMethod]
        public void TokenizeWorksheetNameWithQuotes()
        {
            var input = @"'sheetname'!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(3, tokens.Count());
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorksheetName()
        {
            var input = @"[0]sheetname!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(6, tokens.Count());
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.NameValue));
        }

        [TestMethod]
        public void TokenizeExternalWorksheetNameWithQuotes()
        {
            var input = @"[3]'sheetname'!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(6, tokens.Count());
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorkbookName()
        {
            var input = @"[0]!name";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(5, tokens.Count());
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorkbookInvalidRef()
        {
            var input = @"[0]#Ref!";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(4, tokens.Count());
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.InvalidReference));
        }
        [TestMethod]
        public void TokenizeExternalWorksheetInvalidRef()
        {
            var input = @"[0]Sheet1!#Ref!";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(6, tokens.Count());
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.InvalidReference));
        }

        [TestMethod]
        public void TokenizeShouldHandleWorksheetNameWithSingleQuote()
        {
            var input = @"=VLOOKUP(J7;'Sheet 1''21'!$Q$4:$R$28;2;0)";
            var tokens = _tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(15, tokens.Count());
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.CellAddress));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.CellAddress));
            Assert.IsTrue(tokens[9].TokenTypeIsSet(TokenType.CellAddress));
        }
        [TestMethod]
        public void TokenizeWorksheetAddress()
        {
            var input = @"='Sheet''1'!A1:Name2";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(6, tokens.Count);
        }
        [TestMethod]
        public void TokenizeTableAddress()
        {
            for (int i = 0; i < 100000; i++)
            {
                var input = @"SUM(MyDataTable[[#This Row],[Column 1]])";
                var tokens = _tokenizer.Tokenize(input);
                Assert.AreEqual(13, tokens.Count);
                Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.TableName));
            }
        }
        [TestMethod]
        public void TokenizeTableAddressPerformance()
        {
            var input = @"SUM(MyDataTable[[#This Row],[Column 1]])";
            for (int i = 0; i < 100000; i++)
            {
                var tokens = _tokenizer.Tokenize(input);
            }
        }
        [TestMethod]
        public void TokenizeKeepWhiteSpace()
        {
            var input = @"A1:B3  B2:C5";
            var tokenizer = new OptimizedSourceCodeTokenizer(null, null, false, true);
            var tokens = tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[3].TokenType);
            Assert.AreEqual(8, tokens.Count);
            
            input = "=( A1:B3 )   (B2:C3)";
            tokens = tokenizer.Tokenize(input);
            Assert.AreEqual(15, tokens.Count);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[2].TokenType);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[6].TokenType);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[8].TokenType);
            Assert.AreEqual("   ", tokens[8].Value);

            input = "=( A1:B3 )( B2:C3  )  ";
            tokens = tokenizer.Tokenize( input);
            Assert.AreEqual(16, tokens.Count);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[2].TokenType);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[6].TokenType);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[9].TokenType);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[13].TokenType);
            Assert.AreEqual("  ", tokens[13].Value);
            Assert.AreEqual(TokenType.WhiteSpace, tokens[15].TokenType);
            Assert.AreEqual("  ", tokens[15].Value);
        }
        [TestMethod]
        public void TokenizeWhiteSpace()
        {
            var input = @"A1:B3  B2:C5";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.CellAddress, tokens[2].TokenType);
            Assert.AreEqual(TokenType.Operator, tokens[3].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[4].TokenType);
            Assert.AreEqual(7, tokens.Count);

            input = "=( A1:B3 )   (B2:C3)";
            tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(12, tokens.Count);
            Assert.AreEqual(TokenType.ClosingParenthesis, tokens[5].TokenType);
            Assert.AreEqual(TokenType.Operator, tokens[6].TokenType);
            Assert.AreEqual(TokenType.OpeningParenthesis, tokens[7].TokenType);

            input = "=( A1:B3 )( B2:C3  )  ";
            tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(11, tokens.Count);
            Assert.AreEqual(TokenType.ClosingParenthesis, tokens[5].TokenType);
            Assert.AreEqual(TokenType.OpeningParenthesis, tokens[6].TokenType);
        }
    }
}
