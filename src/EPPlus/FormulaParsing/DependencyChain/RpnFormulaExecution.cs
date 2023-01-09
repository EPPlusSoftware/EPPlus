using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    internal struct CircularReference
    {
        public CircularReference(ulong fromCell, ulong toCell)
        {
            FromCell = fromCell;
            ToCell = toCell;
        }
        internal ulong FromCell;
        internal ulong ToCell;
    }
    internal class RpnOptimizedDependencyChain
    {
        internal List<RpnFormula> _formulas = new List<RpnFormula>();
        internal Stack<RpnFormula> _formulaStack=new Stack<RpnFormula>();
        internal Dictionary<int, RangeDictionary> accessedRanges = new Dictionary<int, RangeDictionary>();
        internal HashSet<ulong> processedCells = new HashSet<ulong>();
        internal List<CircularReference> _circularReferences = new List<CircularReference>();
        internal ISourceCodeTokenizer _tokenizer;
        internal RpnExpressionGraph _graph;
        internal ParsingContext _parsingContext;
        internal RpnFunctionCompilerFactory _functionCompilerFactory;
        public RpnOptimizedDependencyChain(ExcelWorkbook wb, ExcelCalculationOption options)
        {
            _tokenizer = OptimizedSourceCodeTokenizer.Default;
            _parsingContext = wb.FormulaParser.ParsingContext;
            _graph = new RpnExpressionGraph(_parsingContext);

            var parser = wb.FormulaParser;
            var filterInfo = new FilterInfo(wb);
            parser.InitNewCalc(filterInfo);

            _functionCompilerFactory = new RpnFunctionCompilerFactory(_parsingContext.Configuration.FunctionRepository, _parsingContext);
            
            wb.FormulaParser.Configure(config =>
            {
                config.AllowCircularReferences = options.AllowCircularReferences;
                config.PrecisionAndRoundingStrategy = options.PrecisionAndRoundingStrategy;
            });

        }

        internal void Add(RpnFormula f)
        {
            _formulas.Add(f);
        }
        internal RpnOptimizedDependencyChain Execute()
        {
            return RpnFormulaExecution.Execute(_parsingContext.Package.Workbook, new ExcelCalculationOption());
        }
        internal RpnOptimizedDependencyChain Execute(ExcelWorksheet ws)
        {
            return RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
        }
        internal RpnOptimizedDependencyChain Execute(ExcelWorksheet ws, ExcelCalculationOption options)
        {
            return RpnFormulaExecution.Execute(ws, options);
        }
    }
    internal class RpnFormulaExecution
    {
        internal static ArgumentParser _boolArgumentParser = new BoolArgumentParser();
        internal static RpnOptimizedDependencyChain Execute(ExcelWorkbook wb, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(wb, options);
            foreach (var ws in wb.Worksheets)
            {
                if (ws.IsChartSheet==false)
                {
                    ExecuteChain(depChain, ws.Cells, options);
                    ExecuteChain(depChain, ws.Names, options);
                }
            }
            ExecuteChain(depChain, wb.Names, options);

            return depChain;
        }
        internal static RpnOptimizedDependencyChain Execute(ExcelWorksheet ws, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(ws.Workbook, options);

            ExecuteChain(depChain, ws.Cells, options);

            return depChain;
        }
        internal static RpnOptimizedDependencyChain Execute(ExcelRangeBase cells, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(cells._workbook, options);

            ExecuteChain(depChain, cells, options);

            return depChain;
        }
        internal static object ExecuteFormula(ExcelWorksheet ws, string formula, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(ws.Workbook, options);
            return ExecuteChain(depChain, ws, formula, options);
        }
        internal static object ExecuteFormula(ExcelWorkbook wb, string formula, FormulaCellAddress cell, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(wb, options);
            ExcelWorksheet ws;
            if (cell.WorksheetIx < 0 || cell.WorksheetIx >= wb.Worksheets.Count)
            {
                ws = null;
            }
            else
            {
                ws = wb.Worksheets[cell.WorksheetIx];
            }
            return ExecuteChain(depChain, ws, formula, cell, options);
        }
        internal static object ExecuteFormula(ExcelWorkbook wb, string formula, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(wb, options);

            return ExecuteChain(depChain, null, formula, options);
        }
        //internal static object ExecuteFormula(ExcelWorkbook wb, FormulaCellAddress cell, ExcelCalculationOption options)
        //{
        //    var depChain = new RpnOptimizedDependencyChain(wb, options);
        //    var ws = wb.GetWorksheetByIndexInList(cell.WorksheetIx);
        //    if (ws == null) return null;            
        //    return ExecuteChain(depChain, ws, ws.Cells[cell.Row, cell.Column].Formula, options);
        //}


        private static void ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelRangeBase range, ExcelCalculationOption options)
        {
            var ws = range.Worksheet;
            RpnFormula f = null;
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(ws.IndexInList, range._fromRow, range._fromCol);
            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
            while (fs.Next())
            {
                if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (depChain.processedCells.Contains(id) == false)
                {
                    if (GetFormula(depChain, ws, fs, ref f))
                    {
                        AddChainForFormula(depChain, f, options);
                    }
                }
            }
        }
        private static void ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelNamedRangeCollection namesCollection, ExcelCalculationOption options)
        {
            var ws = namesCollection._ws;
            RpnFormula f = null;
            var hasWs = ws != null;
            foreach (ExcelNamedRange name in namesCollection)
            {
                depChain._parsingContext.CurrentCell = new FormulaCellAddress(ws==null ? -1 : ws.IndexInList, name.Index, 0);
                var wsIx = (short)(hasWs ? ws.IndexInList : -1);
                var id = ExcelCellBase.GetCellId(wsIx, name.Index, 0);
                if (depChain.processedCells.Contains(id) == false)
                {
                    if (string.IsNullOrEmpty(name.Formula) == false)
                    {
                        f = GetNameFormula(depChain, ws, depChain._parsingContext.ExcelDataProvider.GetName(0, wsIx, name.Name));
                        AddChainForFormula(depChain, f, options);
                    }
                }
            }
        }
        private static object ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, string formula, FormulaCellAddress cell, ExcelCalculationOption options)
        {
            var f = new RpnFormula(ws, cell.Row, cell.Column);
            f.SetFormula(formula, depChain._tokenizer, depChain._graph);
            return AddChainForFormula(depChain, f, options);
        }

        private static object ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, string formula, ExcelCalculationOption options)
        {
            var f = new RpnFormula(ws, 0, 0);
            f.SetFormula(formula, depChain._tokenizer, depChain._graph);
            f._row = -1;
            return AddChainForFormula(depChain, f, options);
        }
        private static bool GetFormula(RpnOptimizedDependencyChain depChain,  ExcelWorksheet ws, CellStoreEnumerator<object> fs, ref RpnFormula f)
        {
            if (fs.Value is int ix)
            {
                var sf = ws._sharedFormulas[ix];
                f = ws._sharedFormulas[ix].GetRpnFormula(depChain, fs.Row, fs.Column);
            }
            else
            {
                var s = fs.Value.ToString();
                //compiler
                if (string.IsNullOrEmpty(s)) return false;
                f = new RpnFormula(ws, fs.Row, fs.Column);
                f.SetFormula(s, depChain._tokenizer, depChain._graph);
            }
            return true;
        }
        private static RpnFormula GetNameFormula(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, INameInfo name)
        {
            ExcelCellBase.SplitCellId(name.Id, out int wsIx, out int row, out int col);
            var f = new RpnFormula(ws, row , col);
            f.SetFormula(name.Formula, depChain._tokenizer, depChain._graph);
            return f;
        }
        private static object AddChainForFormula(RpnOptimizedDependencyChain depChain, RpnFormula f, ExcelCalculationOption options)
        {
            FormulaRangeAddress address=null;
            RangeDictionary rd = AddAddressToRD(depChain, f._ws==null?-1:f._ws.IndexInList);
            rd?.Merge(f._row, f._column);
        ExecuteFormula:
            var ws = f._ws;
            if (ws == null)
            {
                depChain._parsingContext.CurrentCell = new FormulaCellAddress(0, f._row, 0);
            }
            else
            {
                depChain._parsingContext.CurrentCell = new FormulaCellAddress(f._ws.IndexInList, f._row, f._column);
            }
            if (f._tokenIndex < f._tokens.Count)
            {
                address = ExecuteNextToken(depChain, f);
                if (f._tokenIndex < f._tokens.Count)
                {
                    if (address == null && f._expressions[f._tokenIndex].ExpressionType == ExpressionType.NameValue)
                    {
                        f = GetNameFormula(depChain, ws, ((RpnNamedValueExpression)f._expressions[f._tokenIndex])._name);
                        goto NextFormula;
                    }

                    if (address == null)
                    {
                        address = f._expressions[f._tokenIndex].GetAddress();
                    }
                    if (ws==null)
                    {
                        if (address?.WorksheetIx < 0)
                        {
                            throw (new InvalidOperationException("Address in formula does not reference a worksheet and does not belong to a worksheet."));
                        }
                        else
                        {
                            ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(address.WorksheetIx);
                        }
                    }
                    else if(address?.WorksheetIx >= 0 && ws?.IndexInList != address?.WorksheetIx)
                    {
                        ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(address.WorksheetIx);
                    }

                    rd = AddAddressToRD(depChain, ws.IndexInList);

                    if (rd.Exists(address) || address.CollidesWith(ws.IndexInList, f._row, f._column))
                    {
                        CheckCircularReferences(depChain, f, address, options);
                    }

                    if (rd.Merge(ref address))
                    {
                        goto FollowChain;
                    }
                    f._tokenIndex++;
                    goto ExecuteFormula;
                }
            }
            object value;
            if(f._tokenIndex==int.MaxValue) //int.MaxValue means we have an invalid formulas and we should return a name error 
            {
                value = ExcelErrorValue.Create(eErrorType.Name);
            }
            else
            {
                value = f._expressionStack.Pop().Compile().ResultValue;
            }

            //Set the value.
            if (f._row >= 0)
            {
                if (f._ws == null)
                {
                    depChain._parsingContext.Package.Workbook.Names[f._row].Value = value;
                }
                else
                {
                    if (f._column == 0)
                    {
                        f._ws.Names[f._row].Value = value;
                    }
                    else
                    {
                        f._ws.SetValueInner(f._row, f._column, value);
                    }
                    var id = ExcelCellBase.GetCellId(ws.IndexInList, f._row, f._column);
                    depChain.processedCells.Add(id);
                }
            }
            depChain._formulas.Add(f);
            if (depChain._formulaStack.Count > 0)
            {
                f = depChain._formulaStack.Pop();
                goto NextFormula;
            }
            return value;
        FollowChain:
            ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(address.WorksheetIx);
            f._formulaEnumerator = new CellStoreEnumerator<object>(ws._formulas, address.FromRow, address.FromCol, address.ToRow, address.ToCol);
        NextFormula:
            if (f._formulaEnumerator.Next())
            {
                depChain._formulaStack.Push(f);
                if(GetFormula(depChain, ws, f._formulaEnumerator, ref f))
                {
                    goto ExecuteFormula;
                }
                else
                {
                    goto NextFormula;
                }
               
            }
            f._tokenIndex++;
            goto ExecuteFormula;
        }

        private static RangeDictionary AddAddressToRD(RpnOptimizedDependencyChain depChain, int wsIx)
        {
            if (wsIx < 0) return null;
            if (depChain.accessedRanges.TryGetValue(wsIx, out RangeDictionary rd) == false)
            {
                rd = new RangeDictionary();
                depChain.accessedRanges.Add(wsIx, rd);
            }

            return rd;
        }

        private static void CheckCircularReferences(RpnOptimizedDependencyChain depChain, RpnFormula f, FormulaRangeAddress address, ExcelCalculationOption options)
        {
            if (address.CollidesWith(f._ws.IndexInList, f._row, f._column))
            {
                var fId = ExcelCellBase.GetCellId(f._ws.IndexInList, f._row, f._column);
                HandleCircularReference(depChain, f, options, fId);
            }

            foreach (var sf in depChain._formulaStack)
            {
                var toCell = ExcelCellBase.GetCellId(f._ws.IndexInList, sf._row, sf._column);
                if(address.CollidesWith(f._ws.IndexInList, sf._row, sf._column))
                {
                    HandleCircularReference(depChain, f, options, toCell);
                }
            }
        }

        private static void HandleCircularReference(RpnOptimizedDependencyChain depChain, RpnFormula f, ExcelCalculationOption options, ulong toCell)
        {
            if (options.AllowCircularReferences)
            {
                //var refFormula = depChain._formulaStack.Peek();
                var fromCell = ExcelCellBase.GetCellId(f._ws.IndexInList, f._row, f._column);
                depChain._circularReferences.Add(new CircularReference(fromCell, toCell));
            }
            else
            {
                throw new CircularReferenceException($"Circular reference in cell {f.GetAddress()}");
            }
        }

        private static FormulaRangeAddress ExecuteNextToken(RpnOptimizedDependencyChain depChain, RpnFormula f)
        {

            var s = f._expressionStack;
            while (f._tokenIndex < f._tokens.Count)
            {
                var t = f._tokens[f._tokenIndex];
                switch (t.TokenType)
                {
                    case TokenType.Boolean:
                    case TokenType.Integer:
                    case TokenType.Decimal:
                    case TokenType.StringContent:
                    case TokenType.Array:
                        s.Push(f._expressions[f._tokenIndex]);
                        break;
                    case TokenType.Negator:
                        s.Peek().Negate();
                        break;
                    case TokenType.CellAddress:
                    case TokenType.ExcelAddress:
                        var e = f._expressions[f._tokenIndex];
                        s.Push(e);
                        if (f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek())==false)
                        {
                            return e.GetAddress();
                        }
                        break;
                    case TokenType.NameValue:
                        var ne = (RpnNamedValueExpression)f._expressions[f._tokenIndex];
                        s.Push(ne);
                        if (ne._name != null)
                        {
                            var address = ne.GetAddress();
                            if (address == null)
                            {
                                if (string.IsNullOrEmpty(ne._name?.Formula) == false)
                                {
                                    return null;
                                }
                            }
                            else if (f._funcStack.Count == 0)
                            {
                                return address;
                            }
                        }
                        break;
                    case TokenType.Comma:
                        if(f._funcStack.Count > 0)
                        {
                            var fexp = f._funcStack.Peek();
                            var pi = fexp._function.GetParameterInfo(fexp._argPos++);
                            if (pi == FunctionParameterInformation.Condition)
                            {
                                var v = s.Pop().Compile();
                                PushResult(depChain._parsingContext, f, v);
                                fexp._latestConitionValue = (bool)_boolArgumentParser.Parse(v.Result);
                                f._tokenIndex = GetNextTokenPosFromCondition(f, fexp);
                            }
                            else if (fexp._latestConitionValue.HasValue)
                            {
                                pi = fexp._function.GetParameterInfo(fexp._argPos);
                                if ((pi == FunctionParameterInformation.UseIfConditionIsFalse && fexp._latestConitionValue == true)
                                   ||
                                   (pi == FunctionParameterInformation.UseIfConditionIsTrue && fexp._latestConitionValue == false))
                                {
                                    f._tokenIndex = GetNextTokenPosFromCondition(f, fexp);
                                }
                            }
                            else if(fexp._function.HasNormalArguments)
                            {
                                fexp._arguments.Add(f._tokenIndex);
                            }
                        }
                        break;
                    case TokenType.Function:
                        var r=ExecFunc(depChain, t, f);
                        if(r.DataType==DataType.ExcelRange)
                        {
                            if (f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek()) == false && r.Address!=null)
                            {
                                return r.Address;
                            }
                        }
                        break;
                    case TokenType.StartFunctionArguments:
                        var fe = (RpnFunctionExpression)f._expressions[f._tokenIndex];
                        if(fe._function==null)  //Function does not exists. Push #NAME?
                        {
                            LoadArgumentPositions(fe, f);
                            f._tokenIndex = fe._endPos;
                            f._expressionStack.Push(new RpnErrorExpression(new CompileResult(eErrorType.Name), depChain._parsingContext));
                            break;
                        }
                        if(fe._function.HasNormalArguments==false && fe._arguments.Count <= 1)
                        {
                            LoadArgumentPositions(fe, f);
                        }
                        f._funcStack.Push(fe);
                        break;
                    case TokenType.Operator:
                        ApplyOperator(depChain._parsingContext, t, f);
                        break;
                    case TokenType.Percent:
                        ApplyPercent(depChain._parsingContext, f);
                        break;
                }
                f._tokenIndex++;
            }
            return null;
        }

        private static void ApplyPercent(ParsingContext context, RpnFormula f)
        {
            var e = f._expressionStack.Pop();
            var v=e.Compile().ResultNumeric;
            v /= 100;
            f._expressionStack.Push(new RpnDecimalExpression(new CompileResult(v, DataType.Decimal), context));
        }

        private static bool ShouldIgnoreAddress(RpnFunctionExpression fe)
        {
            return !(fe._function.HasNormalArguments || fe._function.GetParameterInfo(fe._argPos)!=FunctionParameterInformation.IgnoreAddress);
        }

        private static int GetNextTokenPosFromCondition(RpnFormula f, RpnFunctionExpression fexp)
        {
            if(fexp._argPos < fexp._arguments.Count)
            {
                var fe = fexp._function.GetParameterInfo(fexp._argPos);
                while(fexp._argPos < fexp._arguments.Count && (
                    (fe == FunctionParameterInformation.UseIfConditionIsTrue && fexp._latestConitionValue == false) ||
                    (fe == FunctionParameterInformation.UseIfConditionIsFalse && fexp._latestConitionValue == true)
                    ))
                {
                    fexp._argPos++;
                    f._expressionStack.Push(RpnExpression.Empty);  //This expression is not used.
                    fe = fexp._function.GetParameterInfo(fexp._argPos);
                }
                if(fexp._argPos < fexp._arguments.Count)
                {
                    return fexp._arguments[fexp._argPos];
                }
                else
                {
                    return fexp._endPos - 1;
                }
            }
            return f._tokenIndex;
        }

        private static void LoadArgumentPositions(RpnFunctionExpression func, RpnFormula f)
        {
            int subFunctions = 0;
            for(int i=f._tokenIndex+1;i<f._tokens.Count;i++)
            {
                if (f._tokens[i].TokenType==TokenType.Comma)
                {
                    if (subFunctions == 0)
                    {
                        func._arguments.Add(i);
                    }
                }
                else if(f._tokens[i].TokenType==TokenType.StartFunctionArguments)
                {
                    subFunctions++;
                }
                else if (f._tokens[i].TokenType==TokenType.Function)
                {
                    if (subFunctions == 0)
                    {
                        func._endPos = i;
                        return;
                    }
                    subFunctions--;
                }
            }
            func._endPos = f._tokens.Count - 1;
        }

        private static void ApplyOperator(ParsingContext context, Token opToken, RpnFormula f)
        {
            var v1 = f._expressionStack.Pop();
            var v2 = f._expressionStack.Pop();

            var c1 = v1.Compile();
            var c2 = v2.Compile();

            if (OperatorsDict.Instance.TryGetValue(opToken.Value, out IOperator op))
            {
                var result = op.Apply(c2, c1, context);
                PushResult(context, f, result);
            }
        }

        private static CompileResult ExecFunc(RpnOptimizedDependencyChain depChain, Token t, RpnFormula f)
        {
            var func = depChain._parsingContext.Configuration.FunctionRepository.GetFunction(t.Value);
            var args = GetFunctionArguments(f);
            var compiler = depChain._functionCompilerFactory.Create(func);
            CompileResult result;
            try
            {
                result = compiler.Compile(args);
                PushResult(depChain._parsingContext, f, result);
            }
            catch(ExcelErrorValueException e)
            {
                result = new CompileResult(e.ErrorValue, DataType.ExcelError);
                f._expressionStack.Push(new RpnErrorExpression(result, depChain._parsingContext));
            }
            return result;
        }
        private static void PushResult(ParsingContext context, RpnFormula f, CompileResult result)
        {
            switch (result.DataType)
            {
                case DataType.Boolean:
                    f._expressionStack.Push(new RpnBooleanExpression(result, context));
                    break;
                case DataType.Integer:
                    f._expressionStack.Push(new RpnDecimalExpression(result, context));
                    break;
                case DataType.Decimal:
                case DataType.Date:
                case DataType.Time:
                    f._expressionStack.Push(new RpnDecimalExpression(result, context));
                    break;
                case DataType.String:
                case DataType.ExcelAddress:
                    f._expressionStack.Push(new RpnStringExpression(result, context));
                    break;
                case DataType.ExcelError:
                    f._expressionStack.Push(new RpnErrorExpression(result, context));
                    break;
                case DataType.ExcelRange:
                    f._expressionStack.Push(new RpnRangeExpression(result, context, false));
                    break;
                case DataType.Enumerable:
                    f._expressionStack.Push(new RpnEnumerableExpression(result, context));
                    break;

                    //f._expressionStack.Push(new RpnDateExpression(result, context));
                    //break;
                case DataType.Empty:
                    f._expressionStack.Push(RpnExpression.Empty);
                    break;
                default:
                    throw new InvalidOperationException("Unhandled compile result");
            }
        }


        private static IList<RpnExpression> GetFunctionArguments(RpnFormula f)
        {
            var list = new List<RpnExpression>();
            var func = f._funcStack.Pop();
            var s = f._expressionStack;
            for(int i=0;i<func._arguments.Count;i++)
            {
                var si = s.Pop();
                if(si.ExpressionType!=ExpressionType.Empty)
                {
                    si.Status |= RpnExpressionStatus.FunctionArgument;
                }
                list.Insert(0, si);
            }
            return list;
        }

        private static bool GetProcessedAddress(RpnOptimizedDependencyChain depChain, ref FormulaRangeAddress address)
        {
            if (depChain.accessedRanges.TryGetValue(address.WorksheetIx, out RangeDictionary wsRd) == false)
            {
                wsRd = new RangeDictionary();
                depChain.accessedRanges.Add(address.WorksheetIx, wsRd);
            }
            return wsRd.Merge(ref address);
        }
        private static bool GetProcessedAddress(RpnOptimizedDependencyChain depChain, int wsIndex, int row, int col)
        {
            if (depChain.accessedRanges.TryGetValue(wsIndex, out RangeDictionary wsRd) == false)
            {
                wsRd = new RangeDictionary();
                depChain.accessedRanges.Add(wsIndex, wsRd);
            }
            return wsRd.Merge(row, col);
        }
    }
}
