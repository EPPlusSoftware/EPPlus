using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.CellStore;
using System;
using System.Collections.Generic;
using System.IO;
//using OfficeOpenXml.FormulaParsing;
//using OfficeOpenXml.FormulaParsing.ExpressionGraph;
//using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class CalculateCompareDirecory : TestBase
    {
        //private ParsingContext _context;
        private ExcelPackage _package;
        //private ExcelWorksheet _sheet;
        //private ExcelDataProvider _excelDataProvider;
        //private ExpressionGraphBuilder _graphBuilder;
        //private RpnExpressionCompiler _expressionCompiler;

        [TestInitialize]
        public void Initialize()
        {
            
        }

        [TestCleanup]
        public void Cleanup()
        {
        }


        [TestMethod]
        public void VerifyCalculationInCalculateTestDirectory()
        {
            var path = _testInputPathOptional + "CalculationTests\\";
            if(Directory.Exists(path)==false)
            {
                Assert.Inconclusive($"Directory {path} does not exist.");
            }
            foreach(var xlFile in Directory.GetFiles(path, "*.xlsx"))
            {
                VerifyCalculationInPackage(xlFile);
            }
        }

        private void VerifyCalculationInPackage(string xlFile)
        {
            using(var p = new ExcelPackage(xlFile))
            {
                var values = new Dictionary<ulong, object>();
                foreach(var ws in p.Workbook.Worksheets)
                {
                    var cse = new CellStoreEnumerator<object>(ws._formulas);                    
                    foreach(var f in cse)
                    {
                        var id = ExcelCellBase.GetCellId(ws.IndexInList, cse.Row, cse.Column);
                        values.Add(id, ws.GetValue(cse.Row,cse.Column));
                    }
                }
                p.Workbook.ClearFormulaValues();
                p.Workbook.Calculate();

                foreach(var value in values)
                {
                    ExcelCellBase.SplitCellId(value.Key, out int wsIndex, out int row, out int col);
                    var ws = p.Workbook.Worksheets[wsIndex];
                    var v = ws.GetValue(row, col);

                    if (v.Equals(value.Value)==false)
                    {
                        Assert.Fail($"Value differs worksheet {ws.Name} Row {row}, Column  {col}");
                    }
                }
            }
        }
    }
}
