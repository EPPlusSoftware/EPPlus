using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
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
        internal List<RpnFormula> formulas = new List<RpnFormula>();
        internal Stack<RpnFormula> _formulaStack=new Stack<RpnFormula>();
        internal Dictionary<int, RangeDictionary> accessedRanges = new Dictionary<int, RangeDictionary>();
        internal HashSet<ulong> processedCells = new HashSet<ulong>();
        internal List<CircularReference> _circularReferences = new List<CircularReference>();
        internal ISourceCodeTokenizer _tokenizer;
        internal RpnExpressionGraph _graph;
        internal ParsingContext _parsingContext;
        internal RpnFunctionCompilerFactory _functionCompilerFactory;
        public RpnOptimizedDependencyChain(ExcelWorkbook wb)
        {
            _tokenizer = OptimizedSourceCodeTokenizer.Default;
            _parsingContext = ParsingContext.Create(wb._package);
            var dataProvider = new EpplusExcelDataProvider(wb._package, _parsingContext);
            _parsingContext.ExcelDataProvider = dataProvider;
            _parsingContext.NameValueProvider = new EpplusNameValueProvider(dataProvider);
            _parsingContext.RangeAddressFactory = new RangeAddressFactory(dataProvider, _parsingContext);
            _graph = new RpnExpressionGraph(_parsingContext);
            _functionCompilerFactory = new RpnFunctionCompilerFactory(_parsingContext.Configuration.FunctionRepository, _parsingContext);
        }

        internal void Add(RpnFormula f)
        {
            formulas.Add(f);
        }
        internal RpnOptimizedDependencyChain Execute()
        {
            return RpnFormulaExecution.Create(_parsingContext.Package.Workbook, new ExcelCalculationOption());
        }
        internal RpnOptimizedDependencyChain Execute(ExcelWorksheet ws)
        {
            return RpnFormulaExecution.Create(ws, new ExcelCalculationOption());
        }
        internal RpnOptimizedDependencyChain Execute(ExcelWorksheet ws, ExcelCalculationOption options)
        {
            return RpnFormulaExecution.Create(ws, options);
        }
    }
    internal class RpnFormulaExecution
    {
        internal static ArgumentParser _boolArgumentParser = new BoolArgumentParser();
        internal static RpnOptimizedDependencyChain Create(ExcelWorkbook wb, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(wb);
            foreach (var ws in wb.Worksheets)
            {
                if (ws.IsChartSheet==false)
                {
                    ExecuteChain(depChain, wb.FormulaParser.Lexer, ws.Cells, options);
                }
            }
            foreach (var name in wb.Names)
            {
                if (name.NameValue == null)
                {
                    //GetChain(depChain, wb.FormulaParser.Lexer, name, options);
                }
            }
            return depChain;
        }
        internal static RpnOptimizedDependencyChain Create(ExcelWorksheet ws, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(ws.Workbook);

            ExecuteChain(depChain, ws.Workbook.FormulaParser.Lexer, ws.Cells, options);

            return depChain;
        }
        internal static RpnOptimizedDependencyChain Create(ExcelRange cells, ExcelCalculationOption options)
        {
            var depChain = new RpnOptimizedDependencyChain(cells._workbook);

            ExecuteChain(depChain, cells.Worksheet.Workbook.FormulaParser.Lexer, cells, options);

            return depChain;
        }

        private static void ExecuteChain(RpnOptimizedDependencyChain depChain, ILexer lexer, ExcelRange range, ExcelCalculationOption options)
        {
            var ws = range.Worksheet;
            RpnFormula f = null;
            //TODO: Remove the row below when scope has been fixed.
            //depChain._parsingContext.Scopes.NewScope(new FormulaRangeAddress() { FromRow = range._fromRow, FromCol = range._fromCol});
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(ws.IndexInList, range._fromRow, range._fromCol);
            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
            while (fs.Next())
            {
                if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (depChain.processedCells.Contains(id) == false)
                {
                    f=GetFormula(depChain, ws, fs);
                    AddChainForFormula(depChain, lexer, f, options);
                }
            }
        }

        private static RpnFormula GetFormula(RpnOptimizedDependencyChain depChain,  ExcelWorksheet ws, CellStoreEnumerator<object> fs)
        {
            if (fs.Value is int ix)
            {
                var sf = ws._sharedFormulas[ix];
                return ws._sharedFormulas[ix].GetRpnFormula(depChain, fs.Row, fs.Column);
            }
            else
            {
                var s = fs.Value.ToString();
                //compiler
                if (string.IsNullOrEmpty(s)) return null;
                var f = new RpnFormula(ws, fs.Row, fs.Column);
                f.SetFormula(s, depChain._tokenizer, depChain._graph);
                return f;
            }
        }
        private static RpnFormula GetNameFormula(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, INameInfo name)
        {
            ExcelCellBase.SplitCellId(name.Id, out int wsIx, out int row, out int col);
            var f = new RpnFormula(ws, row , col);
            f.SetFormula(name.Formula, depChain._tokenizer, depChain._graph);
            return f;
        }
        private static void AddChainForFormula(RpnOptimizedDependencyChain depChain, ILexer lexer, RpnFormula f, ExcelCalculationOption options)
        {
            FormulaRangeAddress address=null;
            RangeDictionary rd = AddAddressToRD(depChain, f._ws);
            rd.Merge(f._row, f._column);
        ExecuteFormula:
            var ws = f._ws;
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(f._ws.IndexInList, f._row, f._column);
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

                    rd = AddAddressToRD(depChain, ws);
                    if (address == null)
                    {
                        address = f._expressions[f._tokenIndex].GetAddress();
                    }

                    if (rd.Exists(address) || address.CollidesWith(f._ws.IndexInList, f._row, f._column))
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
            //Set the value.
            f._ws.SetValueInner(f._row, f._column, f._expressionStack.Pop().Compile().ResultValue);
            var id = ExcelCellBase.GetCellId(ws.IndexInList, f._row, f._column);
            depChain.processedCells.Add(id);
            depChain.formulas.Add(f);
            if (depChain._formulaStack.Count > 0)
            {
                f = depChain._formulaStack.Pop();
                goto NextFormula;
            }
            return;
        FollowChain:
            ws = depChain._parsingContext.Package.Workbook.Worksheets[address.WorksheetIx];
            f._formulaEnumerator = new CellStoreEnumerator<object>(ws._formulas, address.FromRow, address.FromCol, address.ToRow, address.ToCol);
        NextFormula:
            if (f._formulaEnumerator.Next())
            {
                depChain._formulaStack.Push(f);
                f=GetFormula(depChain, ws, f._formulaEnumerator);
                goto ExecuteFormula;
            }
            f._tokenIndex++;
            goto ExecuteFormula;
        }

        private static RangeDictionary AddAddressToRD(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws)
        {
            if (depChain.accessedRanges.TryGetValue(ws.IndexInList, out RangeDictionary rd) == false)
            {
                rd = new RangeDictionary();
                depChain.accessedRanges.Add(ws.IndexInList, rd);
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
                        var address = ne.GetAddress();
                        if(address==null)
                        {
                            if(string.IsNullOrEmpty(ne._name.Formula)==false)
                            {
                                return null;
                            }
                        }
                        else if (f._funcStack.Count == 0)                        
                        {
                            return address;
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
                            if (f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek()) == false)
                            {
                                return r.Address;
                            }
                        }
                        break;
                    case TokenType.StartFunctionArguments:
                        var fe = (RpnFunctionExpression)f._expressions[f._tokenIndex];
                        if(fe._function.HasNormalArguments==false)
                        {
                            LoadArgumentPositions(fe, f);
                        }
                        f._funcStack.Push(fe);
                        break;
                    case TokenType.Operator:
                        ApplyOperator(depChain._parsingContext, t, f);
                        break;
                }
                f._tokenIndex++;
            }
            return null;
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
            var result = compiler.Compile(args);
            PushResult(depChain._parsingContext, f, result);
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
                    f._expressionStack.Push(new RpnDecimalExpression(result, context));
                    break;
                case DataType.String:
                    f._expressionStack.Push(new RpnStringExpression(result, context));
                    break;
                case DataType.ExcelError:
                    f._expressionStack.Push(new RpnErrorExpression(result, context));
                    break;
                case DataType.ExcelRange:
                    f._expressionStack.Push(new RpnRangeExpression(result.Address, false));
                    break;
                case DataType.Enumerable:
                    f._expressionStack.Push(new RpnEnumerableExpression(null, context));
                    break;
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
