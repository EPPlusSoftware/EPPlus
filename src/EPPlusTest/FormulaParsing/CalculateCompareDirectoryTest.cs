using Microsoft.VisualBasic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
//using OfficeOpenXml.FormulaParsing;
//using OfficeOpenXml.FormulaParsing.FormulaExpressions;
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
        //private ExcelPackage _package;
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
            if (Directory.Exists(path) == false)
            {
                Assert.Inconclusive($"Directory {path} does not exist.");
            }
            
            foreach(var xlFile in Directory.GetFiles(path).Where(x => x.EndsWith(".xlsx") || x.EndsWith(".xlsm")))
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
            var formulaLogFile = new FileInfo("c:\\temp\\formulaLog.log");
            if (formulaLogFile.Exists) formulaLogFile.Delete();
            logWriter.WriteLine($"File {xlFile} starting");
            using(var p = new ExcelPackage(xlFile))
            {
                p.Workbook.FormulaParserManager.AttachLogger(formulaLogFile);
                var values = new Dictionary<ulong, object>();
                foreach (var ws in p.Workbook.Worksheets)
                {                    
                    if (ws.IsChartSheet) continue;
                    var cse = new CellStoreEnumerator<object>(ws._formulas);
                    foreach(var f in cse)
                    {
                        var id = ExcelCellBase.GetCellId(ws.IndexInList, cse.Row, cse.Column);
                        values.Add(id, ws.GetValue(cse.Row, cse.Column));
                    }
                    foreach(var name in ws.Names)
                    {                        
                        var id = ExcelCellBase.GetCellId(ws.IndexInList, name.Index, 0);
                        values.Add(id, name.Value);
                    }
                }

                foreach (var name in p.Workbook.Names)
                {
                    var id = ExcelCellBase.GetCellId(ushort.MaxValue, name.Index, 0);
                    values.Add(id, name.Value);
                }

               // UpdateData(p);
                
                p.Workbook.ClearFormulaValues();
                logWriter.WriteLine($"Calculating {xlFile} starting {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}.  Elapsed {new TimeSpan(sw.ElapsedTicks)}");
                try
                {
                    p.Workbook.Calculate(x => x.CacheExpressions=true);
                    //p.Workbook.Worksheets["Monthly Cash Flow"].Cells["F10"].Calculate(x => x.CacheExpressions = false);
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
                    string nameOrAddress;
                    if (wsIndex ==  ushort.MaxValue || wsIndex==-1)
                    {
                        ws = null;
                        v = p.Workbook.Names[row].Value;
                        nameOrAddress = p.Workbook.Names[row].Name;
                    }
                    else
                    {
                        ws = p.Workbook.Worksheets[wsIndex];
                        if (col == 0)
                        {
                            v = ws.Names[row].Value;
                            nameOrAddress = ws.Names[row].Name;
                        }
                        else
                        { 
                            v = ws.GetValue(row, col);
                            nameOrAddress = ExcelCellBase.GetAddress(row, col);
                        }
                    }

                    //if ((v==null && value.Value!=null) || !(v!=null && v.Equals(value.Value) || ConvertUtil.GetValueDouble(v) == ConvertUtil.GetValueDouble(value.Value)))
                    //{
                    ////Assert.Fail($"Value differs worksheet {ws.Name}\tRow {row}\tColumn  {col}\tDiff");
                    var diff = ConvertUtil.GetValueDouble(v) - ConvertUtil.GetValueDouble(value.Value);
                    //if(col==0)
                    //{
                    //    logWriter.WriteLine($"{ws?.Name}\t{row}\t{value.Value:0.0000000000}\t{v:0.0000000000}\t{diff}");
                    //}
                    //else
                    //{
                    //    logWriter.WriteLine($"{ws?.Name}\t{ExcelCellBase.GetAddress(row, col)}\t{value.Value:0.0000000000}\t{v:0.0000000000}\t{diff}");
                    //}
                    var s1 = (value.Value ?? "").ToString().Replace("\r", "").Replace("\n", "");
                    var s2 = (v ?? "").ToString().Replace("\r", "").Replace("\n", "");
                    
                    logWriter.WriteLine($"{ws?.Name}\t{nameOrAddress}\t{s1}\t{s2}\t{diff}");
                    //}
                }
                logWriter.WriteLine($"File end processing {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}. Elapsed {new TimeSpan(sw.ElapsedTicks).ToString()}");
                logWriter.Close();
                logWriter.Dispose();

                SaveWorkbook("calcIssue.xlsx", p);
            }
        }
        private void UpdateData(ExcelPackage p)
        {
            var inputSheet = p.Workbook.Worksheets["Invoer"];
            if (inputSheet == null) return;
            inputSheet.Cells[2, 1].Value = "Avery 50 gold gloss Polyester op 123 cm";
            inputSheet.Cells[2, 2].Value = 1;
            inputSheet.Cells[2, 3].Value = 0;
            inputSheet.Cells[2, 7].Value = 44910d;
            inputSheet.Cells[2, 8].Value = "Vink VTS";
            inputSheet.Cells[2, 10].Value = "1268";
            inputSheet.Cells[2, 11].Value = "NL";
            inputSheet.Cells[2, 12].Value = "2719JE";
            inputSheet.Cells[2, 13].Value = 17;
            inputSheet.Cells[2, 14].Value = 5592;
            inputSheet.Cells[2, 15].Value = 770347;

            var opties2Sheet = p.Workbook.Worksheets["Opties 2"];
            var outputSheet = p.Workbook.Worksheets["Uitvoer"];
            //opties2Sheet.Calculate();
            //outputSheet.Calculate();
            //package.Workbook.Calculate();
            //outputSheet.Calculate();
            //var outputA2 = outputSheet.Cells[2, 1].Value;
            //var outputB2 = outputSheet.Cells[2, 1].Value;

        }
    }
}
