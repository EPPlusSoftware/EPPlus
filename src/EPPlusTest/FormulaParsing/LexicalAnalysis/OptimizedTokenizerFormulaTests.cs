using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class OptimizedTokenizerFormulaTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("FormulaToken.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Sheet1");
            LoadTestdata(_ws);
            _ws.Tables.Add(_ws.Cells["A1:E101"], "MyTable");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        //[TestMethod]
        //public void VerifyFormulaTokensTable_Performance()
        //{
        //    var r = _ws.Cells["A4:A105"];
        //    for (int i = 0; i < 1000; i++)
        //    {
        //        var f = @"SUM(MyTable[[#This Row],[Date]])";
        //        var formula = new SharedFormula(r, f);
        //    }
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable_ThisRowOneColumn()
        //{
        //    var f = @"SUM(MyTable[[#This Row],[Date]])";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    Assert.AreEqual(13, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.TableName, formula.Tokens[2].TokenType);
        //    Assert.AreEqual(TokenType.TablePart, formula.Tokens[5].TokenType);
        //    Assert.AreEqual(TokenType.TableColumn, formula.Tokens[9].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(4, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(4, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(1, fr.Ranges[0].ToCol);
        //    Assert.IsFalse(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.FromColFixed | FixedFlag.ToColFixed, fr.Ranges[0].FixedFlag);

        //    formula.SetOffset(5, 0);

        //    Assert.AreEqual(9, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(9, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(1, fr.Ranges[0].ToCol);

        //    formula.SetOffset(3, 0);

        //    Assert.AreEqual(7, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(7, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(1, fr.Ranges[0].ToCol);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable_AllOneColumn()
        //{
        //    var f = @"SUM(MyTable[[#all],[Date]])";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    Assert.AreEqual(13, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.TableName, formula.Tokens[2].TokenType);
        //    Assert.AreEqual(TokenType.TablePart, formula.Tokens[5].TokenType);
        //    Assert.AreEqual(TokenType.TableColumn, formula.Tokens[9].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(1, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(101, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(1, fr.Ranges[0].ToCol);
        //    Assert.IsTrue(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.All, fr.Ranges[0].FixedFlag);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable_AllSpanColumns()
        //{
        //    var f = @"SUM(MyTable[[#all],[Date]:[StrValue]])";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    Assert.AreEqual(17, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.TableName, formula.Tokens[2].TokenType);
        //    Assert.AreEqual(TokenType.TablePart, formula.Tokens[5].TokenType);
        //    Assert.AreEqual(TokenType.TableColumn, formula.Tokens[9].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(1, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(101, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(3, fr.Ranges[0].ToCol);
        //    Assert.IsTrue(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.All, fr.Ranges[0].FixedFlag);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable_HeaderDataLastColumn()
        //{
        //    var f = @"SUM(MyTable[[#headers],[#data],[NumFormattedValue]])";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    Assert.AreEqual(17, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.TableName, formula.Tokens[2].TokenType);
        //    Assert.AreEqual(TokenType.TablePart, formula.Tokens[5].TokenType);
        //    Assert.AreEqual(TokenType.TablePart, formula.Tokens[9].TokenType);
        //    Assert.AreEqual(TokenType.TableColumn, formula.Tokens[13].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(1, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(4, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(101, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(4, fr.Ranges[0].ToCol);
        //    Assert.IsTrue(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.All, fr.Ranges[0].FixedFlag);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable_TotalsWithNoTotal()
        //{
        //    //Setup
        //    var f = @"SUM(MyTable[#totals])";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    //Assert
        //    Assert.AreEqual(7, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.TableName, formula.Tokens[2].TokenType);
        //    Assert.AreEqual(TokenType.TablePart, formula.Tokens[4].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(-1, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(-1, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(5, fr.Ranges[0].ToCol);
        //    Assert.IsTrue(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.All, fr.Ranges[0].FixedFlag);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable_Totals()
        //{
        //    //Setup
        //    _ws.Tables[0].ShowTotal = true;
        //    var f = @"SUM(MyTable[#totals])";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    //Assert
        //    Assert.AreEqual(7, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.TableName, formula.Tokens[2].TokenType);
        //    Assert.AreEqual(TokenType.TablePart, formula.Tokens[4].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(102, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(102, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(5, fr.Ranges[0].ToCol);
        //    Assert.IsTrue(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.All, fr.Ranges[0].FixedFlag);

        //    _ws.Tables[0].ShowTotal = false;    //Resore to false to avoid problems with other tests.
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable()
        //{
        //    //Setup
        //    var f = @"SUM(MyTable[])";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    //Assert
        //    Assert.AreEqual(6, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.TableName, formula.Tokens[2].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(2, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(101, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(5, fr.Ranges[0].ToCol);
        //    Assert.IsTrue(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.All, fr.Ranges[0].FixedFlag);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable_AsName()
        //{
        //    //Setup
        //    var f = @"SUM(MyTable)";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    //Assert
        //    Assert.AreEqual(4, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.NameValue, formula.Tokens[2].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(2, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(101, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(5, fr.Ranges[0].ToCol);
        //    Assert.IsTrue(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.All, fr.Ranges[0].FixedFlag);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokensTable_TableName_With_SheetReference()
        //{
        //    //Setup
        //    _ws.Cells["J1"].Formula = "SUM(Sheet1!MyTable)";
        //    var f = @"SUM(Sheet1!MyTable)";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    //Assert
        //    Assert.AreEqual(6, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.NameValue, formula.Tokens[4].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaRange));

        //    var fr = (FormulaRange)formula.TokenInfos[2];
        //    Assert.AreEqual(2, fr.Ranges[0].FromRow);
        //    Assert.AreEqual(1, fr.Ranges[0].FromCol);
        //    Assert.AreEqual(101, fr.Ranges[0].ToRow);
        //    Assert.AreEqual(5, fr.Ranges[0].ToCol);
        //    Assert.IsTrue(fr.IsFixed);
        //    Assert.AreEqual(FixedFlag.All, fr.Ranges[0].FixedFlag);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_FixedCellFormula()
        //{
        //    //Setup
        //    var f = @"SUM($A$1:$C$5)";
        //    var formula = new SharedFormula(_ws.Cells["A4:A105"], f);

        //    //Assert
        //    Assert.AreEqual(6, formula.Tokens.Count);
        //    Assert.AreEqual(2, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.CellAddress, formula.Tokens[2].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaCellAddress));

        //    var c1 = (FormulaCellAddress)formula.TokenInfos[2];
        //    Assert.AreEqual(1, c1.Row);
        //    Assert.AreEqual(1, c1.Col);
        //    Assert.IsTrue(c1.IsFixed);

        //    var c2 = (FormulaCellAddress)formula.TokenInfos[4];
        //    Assert.AreEqual(5, c2.Row);
        //    Assert.AreEqual(3, c2.Col);
        //    Assert.IsTrue(c2.IsFixed);

        //    formula.SetOffset(1, 1);

        //    Assert.AreEqual(1, c1.Row);
        //    Assert.AreEqual(1, c1.Col);

        //    Assert.AreEqual(5, c2.Row);
        //    Assert.AreEqual(3, c2.Col);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_CellFormula1()
        //{
        //    //Setup
        //    var f = @"SUM(A$1:C5)";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);

        //    //Assert
        //    Assert.AreEqual(6, formula.Tokens.Count);
        //    Assert.AreEqual(2, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.CellAddress, formula.Tokens[2].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaCellAddress));

        //    var c1 = (FormulaCellAddress)formula.TokenInfos[2];
        //    Assert.AreEqual(1, c1.Row);
        //    Assert.AreEqual(1, c1.Col);
        //    Assert.IsFalse(c1.IsFixed);

        //    var c2 = (FormulaCellAddress)formula.TokenInfos[4];
        //    Assert.AreEqual(5, c2.Row);
        //    Assert.AreEqual(3, c2.Col);
        //    Assert.IsFalse(c2.IsFixed);

        //    //Offset 1,1
        //    formula.SetOffset(1, 1);

        //    Assert.AreEqual(1, c1.Row);
        //    Assert.AreEqual(2, c1.Col);

        //    Assert.AreEqual(6, c2.Row);
        //    Assert.AreEqual(4, c2.Col);

        //    //Offset 0,5
        //    formula.SetOffset(0, 5);

        //    Assert.AreEqual(1, c1.Row);
        //    Assert.AreEqual(6, c1.Col);

        //    Assert.AreEqual(5, c2.Row);
        //    Assert.AreEqual(8, c2.Col);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_if_returning_range()
        //{
        //    var f = @"SUM(A1:IF(A2=1,A3:A4,B3:B4))";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);
        //    Assert.AreEqual(formula.Tokens.Count, 19);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_CellFormula2()
        //{
        //    //Setup

        //    /*
        //     * Funkar nu fram tills rangen ska returneras som en IRangeInfo. /Mats
        //     */
        //    var f = @"SUM($A1:C$5)";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);

        //    //Assert
        //    Assert.AreEqual(6, formula.Tokens.Count);
        //    Assert.AreEqual(2, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.CellAddress, formula.Tokens[2].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[2], typeof(FormulaCellAddress));

        //    var c1 = (FormulaCellAddress)formula.TokenInfos[2];
        //    Assert.AreEqual(1, c1.Row);
        //    Assert.AreEqual(1, c1.Col);
        //    Assert.IsFalse(c1.IsFixed);

        //    var c2 = (FormulaCellAddress)formula.TokenInfos[4];
        //    Assert.AreEqual(5, c2.Row);
        //    Assert.AreEqual(3, c2.Col);
        //    Assert.IsFalse(c2.IsFixed);

        //    //Offset 1,1
        //    formula.SetOffset(1, 1);

        //    Assert.AreEqual(2, c1.Row);
        //    Assert.AreEqual(1, c1.Col);

        //    Assert.AreEqual(5, c2.Row);
        //    Assert.AreEqual(4, c2.Col);

        //    //Offset 0,5
        //    formula.SetOffset(0, 5);

        //    Assert.AreEqual(1, c1.Row);
        //    Assert.AreEqual(1, c1.Col);

        //    Assert.AreEqual(5, c2.Row);
        //    Assert.AreEqual(8, c2.Col);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_NameValue()
        //{
        //    //Setup
        //    _ws.Workbook.Names.AddValue("NameValue", 73);
        //    var f = @"=NameValue";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);

        //    //Assert
        //    Assert.AreEqual(2, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.NameValue, formula.Tokens[1].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[1], typeof(FormulaFixedValue));

        //    var fv = (FormulaFixedValue)formula.TokenInfos[1];
        //    Assert.IsTrue(fv.IsFixed);
        //    Assert.AreEqual(73, fv.Value);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_NameFormula()
        //{
        //    //Setup
        //    var expectedFormula = "A1+B1";
        //    _ws.Workbook.Names.AddFormula("NameFormula", expectedFormula);
        //    var f = @"NameFormula";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);

        //    //Assert
        //    Assert.AreEqual(1, formula.Tokens.Count);
        //    Assert.AreEqual(1, formula.TokenInfos.Count);

        //    Assert.AreEqual(TokenType.NameValue, formula.Tokens[0].TokenType);
        //    Assert.IsInstanceOfType(formula.TokenInfos[0], typeof(FormulaNamedFormula));

        //    var fv = (FormulaNamedFormula)formula.TokenInfos[0];

        //    Assert.IsFalse(fv.IsFixed);
        //    Assert.AreEqual(expectedFormula, fv.Formula);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_MultiAdresses()
        //{
        //    var f = @"SUM($A1:C$5:B12)";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_SheetAdresses()
        //{
        //    var f = @"SUM('sheet1'!$A1:C$5:B12)";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_OffsetFirst()
        //{
        //    var f = "SUM(OFFSET(A3, -1, 0):A1:OFFSET(A3, 1, 0))";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);
        //}
        //[TestMethod]
        //public void VerifyFormulaTokens_ExternalReference()
        //{
        //    var f = "SUM([1]'Sheet1'!A5:A8)";
        //    var formula = new SharedFormula(_ws.Cells["D4:E12"], f);
        //}
    }
}