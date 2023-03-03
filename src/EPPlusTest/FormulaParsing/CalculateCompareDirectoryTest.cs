using Microsoft.VisualBasic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
            if (Directory.Exists(path)==false)
            {
                Assert.Inconclusive($"Directory {path} does not exist.");
            }
            foreach(var xlFile in Directory.GetFiles(path).Where(x=>x.EndsWith(".xlsx") || x.EndsWith(".xlsm")))
            {
                string logFile = path + new FileInfo(xlFile).Name + ".log";
                VerifyCalculationInPackage(xlFile, logFile);
            }
        }

        private void VerifyCalculationInPackage(string xlFile, string logFile)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();            
            if(File.Exists(logFile))
            {
                File.Delete(logFile);
            }
            var logWriter = new StreamWriter(File.OpenWrite(logFile));
            logWriter.WriteLine($"File {xlFile} starting");
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
                logWriter.WriteLine($"Calculating {xlFile} starting. Elapsed {new TimeSpan(sw.ElapsedTicks).ToString()}");
                try
                {
                    p.Workbook.Calculate();
                    //p.Workbook.Worksheets["ERRP"].Cells["Q176"].Calculate();    // 1 891 446
                    //p.Workbook.Worksheets["UAP SUMMARY"].Cells["O10"].Calculate();
                    //p.Workbook.Worksheets["T-UAP"].Cells["B10:B11"].Calculate();
                    //p.Workbook.Worksheets["T-UAP"].Cells["B1"].Calculate();
                    //p.Workbook.Worksheets["ERRP"].Cells["Q198"].Calculate();
                    //p.Workbook.Worksheets["T-Input"].Cells["Q670"].Calculate();
                }
                catch (Exception ex)
                {
                    logWriter.WriteLine($"An exception occured: {ex}");
                }
                logWriter.WriteLine($"Calculating {xlFile} end. Elapsed {new TimeSpan(sw.ElapsedTicks)}");
                logWriter.WriteLine($"Differences:");
                logWriter.WriteLine($"Worksheet\tRow\tColumn\tValue Excel\tValue EPPlus");
                foreach (var value in values)
                {
                    ExcelCellBase.SplitCellId(value.Key, out int wsIndex, out int row, out int col);
                    var ws = p.Workbook.Worksheets[wsIndex];
                    var v = ws.GetValue(row, col);

                    if ((v==null && value.Value!=null) || !(v.Equals(value.Value) || ConvertUtil.GetValueDouble(v) == ConvertUtil.GetValueDouble(value.Value)))
                    {
                        //Assert.Fail($"Value differs worksheet {ws.Name} Row {row}, Column  {col}");
                        logWriter.WriteLine($"{ws.Name}\t{ExcelCellBase.GetAddress(row,col)}\t{value.Value:0.00000000}\t{v:0.00000000}");
                    }
                }
                logWriter.WriteLine($"File end processing. Elapsed {new TimeSpan(sw.ElapsedTicks).ToString()}");
                logWriter.Close();
                logWriter.Dispose();
            }
        }
    }
}
