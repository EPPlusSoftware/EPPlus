using Microsoft.VisualBasic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
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
                    foreach(var name in ws.Names)
                    {
                        var id = ExcelCellBase.GetCellId(ws.IndexInList, name.Index, 0);
                        values.Add(id, name.Value);
                    }
                }
                foreach (var name in p.Workbook.Names)
                {
                    var id = ExcelCellBase.GetCellId(-1, name.Index, 0);
                    values.Add(id, name.Value);
                }

                p.Workbook.ClearFormulaValues();
                logWriter.WriteLine($"Calculating {xlFile} starting {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.  Elapsed {new TimeSpan(sw.ElapsedTicks)}");
                try
                {
                    //p.Workbook.Calculate();
                    //p.Workbook.Worksheets["CELP"].Cells["O128"].Calculate();
                    p.Workbook.Worksheets["CELP"].Cells["N9"].Calculate();
                    //p.Workbook.Worksheets["Stacked Logs"].Cells["N3"].Calculate();  //#REF! not hanlded in lookup
                    //p.Workbook.Names[0].Calculate();  //#REF! not hanlded in lookup
                    //p.Workbook.Worksheets["T-UAP"].Cells["M3"].Calculate();
                    //p.Workbook.Worksheets["T-UAP"].Cells["F66"].Calculate();
                    //p.Workbook.Worksheets["UAP Summary"].Cells["J53"].Calculate(); //#Ref! in And
                    //p.Workbook.Worksheets["ERRP"].Cells["K723"].Calculate();
                    //p.Workbook.Worksheets["ERRP"].Cells["Q176"].Calculate();   
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
                logWriter.WriteLine($"Formula values to compare: {values.Count}");
                logWriter.WriteLine($"Worksheet\tCell\tValue Excel\tValue EPPlus");
                foreach (var value in values)
                {
                    ExcelCellBase.SplitCellId(value.Key, out int wsIndex, out int row, out int col);
                    object v;
                    ExcelWorksheet ws;
                    if (wsIndex < 0)
                    {
                        ws = null;
                        v = p.Workbook.Names[row].Value;
                    }
                    else
                    {
                        ws = p.Workbook.Worksheets[wsIndex];
                        if (col == 0)
                        {
                            v = p.Workbook.Names[row].Value;
                        }
                        else
                        { 
                            v = ws.GetValue(row, col);
                        }
                    }

                    if ((v==null && value.Value!=null) || !(v.Equals(value.Value) || ConvertUtil.GetValueDouble(v) == ConvertUtil.GetValueDouble(value.Value)))
                    {
                        //Assert.Fail($"Value differs worksheet {ws.Name}\tRow {row}\tColumn  {col}\tDiff");
                        var diff = ConvertUtil.GetValueDouble(v) - ConvertUtil.GetValueDouble(value.Value);
                        if(col==0)
                        {
                            logWriter.WriteLine($"{ws?.Name}\t{row}\t{value.Value:0.0000000000}\t{v:0.0000000000}\t{diff}");
                        }
                        else
                        {
                            logWriter.WriteLine($"{ws?.Name}\t{ExcelCellBase.GetAddress(row, col)}\t{value.Value:0.0000000000}\t{v:0.0000000000}\t{diff}");
                        }
                    }
                }
                logWriter.WriteLine($"File end processing {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}. Elapsed {new TimeSpan(sw.ElapsedTicks).ToString()}");
                logWriter.Close();
                logWriter.Dispose();
            }
        }
    }
}
