//using OfficeOpenXml.Core.CellStore;
//using OfficeOpenXml.FormulaParsing.Excel.Functions;
//using OfficeOpenXml.FormulaParsing.Exceptions;
//using OfficeOpenXml.FormulaParsing.FormulaExpressions;
//using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
//using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;

//namespace OfficeOpenXml.FormulaParsing
//{
//    internal class OptimizedDependencyChain
//    {
//        internal List<Formula> formulas = new List<Formula>();
//        internal Dictionary<int, RangeHashset> accessedRanges = new Dictionary<int, RangeHashset>();
//        internal HashSet<ulong> processedCells = new HashSet<ulong>();
//        internal List<ulong> _circularReferences = new List<ulong>();
//        internal void Add(Formula f)
//        {
//            formulas.Add(f);
//        }
//    }
//    internal class OptimizedDependencyChainFactory
//    {
//        internal static OptimizedDependencyChain Create(ExcelWorkbook wb, ExcelCalculationOption options)
//        {
//            var depChain = new OptimizedDependencyChain();
//            foreach (var ws in wb.Worksheets)
//            {
//                if (ws.IsChartSheet==false)
//                {
//                    AddRangeToChain(depChain, wb.FormulaParser.Lexer, ws.Cells, options);
//                    //GetWorksheetNames(ws, depChain, options);
//                }
//            }
//            foreach (var name in wb.Names)
//            {
//                if (name.NameValue == null)
//                {
//                    //GetChain(depChain, wb.FormulaParser.Lexer, name, options);
//                }
//            }
//            return depChain;
//        }
//        internal static OptimizedDependencyChain Create(ExcelWorksheet ws, ExcelCalculationOption options)
//        {
//            var depChain = new OptimizedDependencyChain();

//            AddRangeToChain(depChain, ws.Workbook.FormulaParser.Lexer, ws.Cells, options);

//            return depChain;
//        }
//        internal static OptimizedDependencyChain Create(ExcelRange cells, ExcelCalculationOption options)
//        {
//            var depChain = new OptimizedDependencyChain();

//            AddRangeToChain(depChain, cells.Worksheet.Workbook.FormulaParser.Lexer, cells, options);

//            return depChain;
//        }

//        private static void AddRangeToChain(OptimizedDependencyChain depChain, ILexer lexer, ExcelRange range, ExcelCalculationOption options)
//        {
//            var ws = range.Worksheet;
//            Formula f = null;
//            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
//            while (fs.Next())
//            {

//                if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
//                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
//                if (depChain.processedCells.Contains(id) == false)
//                {
//                    depChain.processedCells.Add(id);
//                    ws.Workbook.FormulaParser.ParsingContext.CurrentCell = new FormulaCellAddress(ws.IndexInList, fs.Row, fs.Column);
//                    if (fs.Value is int ix)
//                    {
//                        f = ws._sharedFormulas[ix].GetFormula(fs.Row, fs.Column);
//                    }

//                    else
//                    {
//                        var s = fs.Value.ToString();
//                        //compiler
//                        if (string.IsNullOrEmpty(s)) continue;
//                        f = new Formula(ws, fs.Row, fs.Column, s);
//                    }
//                    AddChainForFormula(depChain, lexer, f, options);
//                }
//            }
//        }
//        internal class CalcState
//        {
//            internal Stack<Formula> _stack = new Stack<Formula>();

//        }
//        private static void AddChainForFormula(OptimizedDependencyChain depChain, ILexer lexer, Formula f, ExcelCalculationOption options)
//        {
//            var subCalcs = new Stack<CalcState>();
//            var calcState = new CalcState();
//            var ws = f._ws;
//            ExcelFunction currentFunction = null;
////FollowFormulaChain:
//            //var et = f.ExpressionTree;
////            if (f.AddressExpressionIndex < et.AddressExpressions.Count)
////            {
////                var ae = et.AddressExpressions[f.AddressExpressionIndex++];
////                if (ae.ExpressionType == ExpressionType.Function) goto FollowFormulaChain;
////                if(ae._parent?.ExpressionType==ExpressionType.Function)
////                {
////                    var fe = ((FunctionExpression)ae._parent);
////                    currentFunction = fe.Function;
////                    switch(currentFunction.GetParameterInfo(fe.GetArgumentIndex(ae)))
////                    {
////                        case FunctionParameterInformation.IgnoreAddress:
////                            goto FollowFormulaChain;
////                        case FunctionParameterInformation.Condition:
////                            subCalcs.Push(calcState);
////                            calcState = new CalcState();
                            
////                            goto FollowFormulaChain;
////                        default:
////                            break;
////                    }
////                    if(currentFunction.ReturnsReference)
////                    {
////                        //fa.Stack
////                        int i=1;
////                        //var compiler = new FunctionCompilerFactory();
////                    }
////                }
////                var address = ae.Compile().Address;                
////                if (address.FromRow == address.ToRow && address.FromCol == address.ToCol)
////                {
////                    if (GetProcessedAddress(depChain, (int)address.WorksheetIx, address.FromRow, address.FromCol))                         
////                    {
////                        ExcelWorksheet fws;
////                        if (address.WorksheetIx > 0)
////                            fws = ws.Workbook.Worksheets[address.WorksheetIx];
////                        else
////                            fws = ws;

////                        if(fws._formulas.Exists(address.FromRow, address.FromCol))
////                        {
////                            calcState._stack.Push(f);
////                            var fv = fws._formulas.GetValue(address.FromRow, address.FromCol);
////                            if (fv is int ix)
////                            {
////                                f = fws._sharedFormulas[ix].GetFormula(address.FromRow, address.FromCol);
////                            }
////                            else
////                            {
////                                var s = fv.ToString();
////                                //compiler
////                                if (string.IsNullOrEmpty(s)) goto FollowFormulaChain;
////                                f = new Formula(fws, address.FromRow, address.FromCol, s);
////                            }
////                            depChain.processedCells.Add(f.Id);
////                            ws = fws;
////                            goto FollowFormulaChain;
////                        }
////                    }
////                }
////                else if (GetProcessedAddress(depChain, ref address))
////                {
////                    ExcelWorksheet fws;
////                    if (address.WorksheetIx > 0)
////                        fws = ws.Workbook.Worksheets[address.WorksheetIx];
////                    else
////                        fws = ws;

////                    f._formulaEnumerator = new CellStoreEnumerator<object>(fws._formulas, address.FromRow, address.FromCol, address.ToRow, address.ToCol);
////                    goto NextFormula;
////                }
////                if (f.AddressExpressionIndex < et.AddressExpressions.Count)
////                {
////                    //f.AddressExpressionIndex++;
////                    goto FollowFormulaChain;
////                }
////            }
////            if (IsCircularReference(depChain, calcState._stack, f.Id))
////            {
////                //Check
////            }
////            else
////            {
////                depChain.Add(f);
////            }

////            if (calcState._stack.Count > 0)
////            {
////                f = calcState._stack.Pop();
////                ws = f._ws;
////                if (f._formulaEnumerator == null)
////                {
////                    goto FollowFormulaChain;
////                }
////                else
////                {
////                    goto NextFormula;
////                }
////            }
////            return;
////NextFormula:
////            var fs = f._formulaEnumerator;
////            if (f._formulaEnumerator.Next())
////            {
////                if (fs.Value == null || fs.Value.ToString().Trim() == "") goto NextFormula;
////                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
////                if (depChain.processedCells.Contains(id) == false)
////                {
////                    depChain.processedCells.Add(id);
////                    ws.Workbook.FormulaParser.ParsingContext.CurrentCell = new FormulaCellAddress(ws.IndexInList, fs.Row, fs.Column);
////                    calcState._stack.Push(f);
////                    if (fs.Value is int ix)
////                    {
////                        f = ws._sharedFormulas[ix].GetFormula(fs.Row, fs.Column);
////                    }

////                    else
////                    {
////                        var s = fs.Value.ToString();
////                        //compiler
////                        if (string.IsNullOrEmpty(s)) goto NextFormula;
////                        f = new Formula(ws, fs.Row, fs.Column, s);
////                    }
////                    ws = f._ws;
////                    goto FollowFormulaChain;
////                }
////                else if (IsCircularReference(depChain, calcState._stack, id))
////                {
////                    //Check
////                }

////                goto NextFormula;
////            }
////            f._formulaEnumerator = null;
////            goto FollowFormulaChain;
//        }

//        private static bool IsCircularReference(OptimizedDependencyChain depChain, Stack<Formula> stack, ulong Id)
//        {
//            foreach(var f in stack)
//            {
//                var fId = ExcelCellBase.GetCellId(f._ws.IndexInList, f.StartRow, f.StartCol);
//                if (Id==fId)
//                {
//                    depChain._circularReferences.Add(Id);
//                    //throw Circual Reference.
//                    //throw new CircularReferenceException($"Circular reference detected in cell {ExcelCellBase.GetAddress(f.StartRow,f.StartCol)}");
//                    return true;
//                }
//            }
//            return false;
//        }

//        private static bool GetProcessedAddress(OptimizedDependencyChain depChain, ref FormulaRangeAddress address)
//        {
//            if (depChain.accessedRanges.TryGetValue(address.WorksheetIx, out RangeHashset wsRd) == false)
//            {
//                wsRd = new RangeHashset();
//                depChain.accessedRanges.Add(address.WorksheetIx, wsRd);
//            }
//            return wsRd.Merge(ref address);
//        }
//        private static bool GetProcessedAddress(OptimizedDependencyChain depChain, int wsIndex, int row, int col)
//        {
//            if (depChain.accessedRanges.TryGetValue(wsIndex, out RangeHashset wsRd) == false)
//            {
//                wsRd = new RangeHashset();
//                depChain.accessedRanges.Add(wsIndex, wsRd);
//            }
//            return wsRd.Merge(row, col);
//        }
//    }
//}
