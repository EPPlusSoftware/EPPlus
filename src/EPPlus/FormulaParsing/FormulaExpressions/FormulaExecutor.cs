/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2022         EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class FormulaExecutor
    {
        private ParsingContext _parsingContext;
        private FunctionCompilerFactory _functionCompilerFactory;

        internal FormulaExecutor(ParsingContext parsingContext)
        {
            _parsingContext = parsingContext;
            _functionCompilerFactory = new FunctionCompilerFactory(_parsingContext.Configuration.FunctionRepository, _parsingContext);
        }

        internal static List<Token> CreateRPNTokens(IList<Token> tokens)
        {
            var operators = OperatorsDict.Instance;
            Stack<Token> operatorStack = new Stack<Token>();
            var expressions = new List<Token>();
            for (int i = 0; i < tokens.Count; i++)
            {
                Token token = tokens[i];
                switch (token.TokenType)
                {
                    case TokenType.OpeningParenthesis:
                        operatorStack.Push(token);
                        break;
                    case TokenType.ClosingParenthesis:
                        if (operatorStack.Count > 0)
                        {
                            
                            var o = operatorStack.Pop();
                            while (o.TokenType != TokenType.OpeningParenthesis)
                            {
                                expressions.Add(o);
                                if (operatorStack.Count == 0) throw new InvalidOperationException("No closing parenthesis");
                                o = operatorStack.Pop();
                            }
                            if (operatorStack.Count > 0 && operatorStack.Peek().TokenType == TokenType.Function)
                            {
                                expressions.Add(operatorStack.Pop());
                            }
                        }
                        break;
                    case TokenType.Operator:
                    case TokenType.Negator:
                        if (operatorStack.Count > 0)
                        {
                            var o2 = operatorStack.Peek();
                            while (o2.TokenType == TokenType.Operator &&
                                operators[o2.Value].Precedence <= operators[token.Value].Precedence && token.TokenType != TokenType.Negator)
                            {
                                expressions.Add(operatorStack.Pop());
                                if (operatorStack.Count == 0) break;
                                o2 = operatorStack.Peek();
                            }
                        }
                        operatorStack.Push(token);
                        break;

                    case TokenType.Function:
                        expressions.Add(new Token(token.Value,TokenType.StartFunctionArguments));
                        operatorStack.Push(token);
                        break;
                    case TokenType.Comma:
                        if(operatorStack.Count > 0 && tokens[i-1].TokenType != TokenType.ClosingBracket) //If inside a table 
                        {
                            var op = operatorStack.Peek().TokenType;
                            while (op == TokenType.Operator || op == TokenType.Negator)
                            {
                                expressions.Add(operatorStack.Pop());
                                if(operatorStack.Count == 0) break;
                                op = operatorStack.Peek().TokenType;
                            }
                        }
                        expressions.Add(token);
                        break;
                    default:
                        expressions.Add(token);
                        break;
                }

            }

            while (operatorStack.Count > 0)
            {
                expressions.Add(operatorStack.Pop());
            }

            return expressions;
        }
        public static Dictionary<int, Expression> CompileExpressions(ref IList<Token> tokens, ParsingContext parsingContext)
        {
            short extRefIx = short.MinValue;
            int wsIx = int.MinValue;
            var stack = new Stack<FunctionExpression>();
            var expressions = new Dictionary<int, Expression>();
            for (int i = 0; i < tokens.Count; i++)
            {
                var t = tokens[i];
                switch (t.TokenType)
                {
                    case TokenType.Boolean:
                        expressions.Add(i, new BooleanExpression(t.Value, parsingContext));
                        break;
                    case TokenType.Integer:
                        expressions.Add(i, new IntegerExpression(t.Value, parsingContext));
                        break;
                    case TokenType.Decimal:
                        expressions.Add(i, new DecimalExpression(t.Value, parsingContext));
                        break;
                    case TokenType.StringContent:
                        expressions.Add(i, new StringExpression(t.Value, parsingContext));
                        break;
                    case TokenType.CellAddress:
                    case TokenType.FullColumnAddress:
                    case TokenType.FullRowAddress:
                        if (i > 1 && tokens[i - 1].TokenTypeIsAddress && tokens[i + 1].Value == ":" && tokens[i + 1].TokenType == TokenType.Operator)
                        {
                            //We have a two cell addresses with with a colon. Remove tokens and replace with full column address, for example A1:C2.
                            var e = expressions[i - 1];
                            e.MergeAddress(t.Value);
                            tokens.RemoveAt(i - 1);
                            tokens.RemoveAt(i);
                            i--;
                            tokens[i] = new Token(e.GetAddress().WorksheetAddress, TokenType.ExcelAddress);
                        }
                        else
                        {
                            expressions.Add(i, new RangeExpression(t.Value, parsingContext, extRefIx, wsIx));
                        }
                        extRefIx = short.MinValue;
                        wsIx = int.MinValue;
                        break;
                    case TokenType.NameValue:
                        expressions.Add(i, new NamedValueExpression(t.Value, parsingContext, extRefIx, wsIx));
                        break;
                    case TokenType.ExternalReference:
                        if (t.Value.All(c => c >= '0' && c <= '9'))
                        {
                            extRefIx = short.Parse(t.Value);
                        }
                        else
                        {
                            extRefIx = (short)(parsingContext.Package.Workbook.ExternalLinks.GetExternalLink(t.Value)+1);
                        }
                        wsIx = int.MinValue;
                        break;
                    case TokenType.WorksheetNameContent:
                        if (extRefIx <= 0)
                        {
                            wsIx = parsingContext.Package.Workbook.Worksheets.GetPositionByToken(t.Value);
                        }
                        else
                        {
                            wsIx = parsingContext.Package.Workbook.ExternalLinks.GetPositionByToken(extRefIx, t.Value);
                        }
                        break;
                    case TokenType.TableName:
                        ExtractTableAddress(extRefIx, tokens, i, out FormulaTableAddress tableAddress, parsingContext);
                        
                        expressions.Add(i, new TableAddressExpression(tableAddress, parsingContext));
                        break;
                    case TokenType.OpeningEnumerable:
                        ExtractArray(tokens, i , out IRangeInfo rangInfo, parsingContext);
                        expressions.Add(i, new EnumerableExpression(rangInfo, parsingContext));
                        break;
                    case TokenType.StartFunctionArguments:
                        var func = new FunctionExpression(t.Value, parsingContext, i);
                        expressions.Add(i, func);
                        if(i <= tokens.Count && tokens[i+1].TokenType != TokenType.Function) // Check that the function has any argument
                        {
                            func._arguments.Add(i);
                        }
                        stack.Push(func);
                        break;
                    case TokenType.Comma:
                        if (stack.Count > 0)
                        {
                            stack.Peek()._arguments.Add(i);
                        }
                        break;
                    case TokenType.Function:
                        var f = stack.Pop();
                        f._endPos= i;
                        break;
                    case TokenType.InvalidReference:
                        expressions.Add(i, ErrorExpression.RefError);
                        break;
                }
            }
            return expressions;
        }
        //private static void LoadArgumentPositions(FunctionExpression func, IList<Token> tokens, int tokenIndex)
        //{
        //    int subFunctions = 0;
        //    for (int i = tokenIndex + 1; i < tokens.Count; i++)
        //    {
        //        if (tokens[i].TokenType == TokenType.Comma)
        //        {
        //            if (subFunctions == 0)
        //            {
        //                func._arguments.Add(i);
        //            }
        //        }
        //        else if (tokens[i].TokenType == TokenType.StartFunctionArguments)
        //        {
        //            subFunctions++;
        //        }
        //        else if (tokens[i].TokenType == TokenType.Function)
        //        {
        //            if (subFunctions == 0)
        //            {
        //                func._endPos = i;
        //                return;
        //            }
        //            subFunctions--;
        //        }
        //    }
        //    func._endPos = tokens.Count - 1;
        //}
        //Dictionary<int, RangeHashset> _usedRanges;

        //internal CompileResult Execute(IList<Token> exps)
        //{
        //    _usedRanges = new Dictionary<int, RangeHashset>();
        //    var cell = new FormulaCell();
        //    short extRefIx = short.MinValue;
        //    int wsIx = int.MinValue;
        //    var s = cell._expressionStack;

        //    for (int i = 0; i < exps.Count; i++)
        //    {
        //        var t = exps[i];

        //        if (s.Count > 0 && 
        //            !(t.TokenType == TokenType.Operator && t.Value != ":") && 
        //            s.Peek().Status == ExpressionStatus.IsAddress)
        //        {
        //            //We have an address, follow dependency chain before executing .
        //            var a = GetAddressToFollow(s.Peek());
        //            if(a!=null)
        //            {

        //            }
        //        }

        //        switch (t.TokenType)
        //        {                    
        //            case TokenType.Boolean:
        //                s.Push(new BooleanExpression(t.Value, _parsingContext));
        //                break;
        //            case TokenType.Integer:
        //                s.Push(new IntegerExpression(t.Value, _parsingContext));
        //                break;
        //            case TokenType.Decimal:
        //                s.Push(new DecimalExpression(t.Value, _parsingContext));
        //                break;
        //            case TokenType.StringContent:
        //                s.Push(new StringExpression(t.Value, _parsingContext));
        //                break;                    
        //            case TokenType.Negator:
        //                s.Peek().Negate();
        //                break;
        //            case TokenType.CellAddress:
        //                s.Push(new RangeExpression(t.Value, _parsingContext, extRefIx, wsIx));
        //                extRefIx = short.MinValue;
        //                wsIx = int.MinValue;                        
        //                break;
        //            case TokenType.NameValue:
        //                s.Push(new NamedValueExpression(t.Value, _parsingContext, extRefIx, wsIx));
        //                break;
        //            case TokenType.ExternalReference:
        //                extRefIx = short.Parse(t.Value);
        //                break;
        //            case TokenType.WorksheetNameContent:
        //                wsIx = _parsingContext.Package.Workbook.Worksheets.GetPositionByToken(t.Value);
        //                break;
        //            case TokenType.Comma:
        //                cell._funcStackPosition.Peek()._arguments.Add(i-1);
        //                break;
        //            case TokenType.Function:
        //                ExecFunc(t, cell);
        //                break;
        //            case TokenType.StartFunctionArguments:
        //                var func = new FunctionExpression(t.Value, _parsingContext, i);
        //                if (i <= exps.Count && exps[i + 1].TokenType != TokenType.Function)
        //                {
        //                    func._arguments.Add(i);
        //                }
        //                break;
        //            case TokenType.TableName:
        //                ExtractTableAddress(extRefIx, exps, i, out FormulaTableAddress tableAddress, _parsingContext);
        //                s.Push(new TableAddressExpression(tableAddress, _parsingContext));
        //                break;
        //            case TokenType.OpeningEnumerable:
        //                ExtractArray(exps, i, out IRangeInfo range, _parsingContext);
        //                s.Push(new EnumerableExpression(range, _parsingContext));
        //                break;
        //            case TokenType.Operator:
        //                ApplyOperator(t, cell);
        //                break;
        //        }
        //    }
        //    return s.Pop().Compile();
        //}

        //private FormulaRangeAddress GetAddressToFollow(Expression ae)
        //{
        //    var a = ae.Compile().Address;
        //    if (a.WorksheetIx < 0) return null;

        //    RangeHashset rd;
        //    if (!_usedRanges.TryGetValue(a.WorksheetIx, out rd))
        //    {
        //        rd = new RangeHashset();
        //        _usedRanges.Add(a.WorksheetIx, rd);
        //    }

        //    if (a.IsSingleCell)
        //    {
        //        if (rd.Exists(a.FromRow, a.ToRow))
        //        {
        //            return null;
        //        }
        //        var ws = _parsingContext.Package.Workbook.Worksheets[a.WorksheetIx];
        //        if(ws._formulas.Exists(a.FromRow, a.FromCol))
        //        {
        //            return a.Address;
        //        }
        //    }
        //    else
        //    {
        //        FormulaRangeAddress r=a.Address;
        //        if (rd.Merge(ref r))
        //        {

        //        }
        //    }
        //    return null;
        //}

        private static void ExtractTableAddress(int extRef, IList<Token> exps, int i, out FormulaTableAddress tableAddress, ParsingContext parsingContext)
        {
            //var adr = exps[i].Value;
            tableAddress = new FormulaTableAddress(parsingContext) {ExternalReferenceIx = extRef, TableName = exps[i].Value };
            exps.RemoveAt(i);
            int bracketCount=0;
            while (i < exps.Count)
            {
                var t = exps[i];
                switch(t.TokenType)
                {
                    case TokenType.OpeningBracket:
                        bracketCount++;
                        break;
                    case TokenType.ClosingBracket:
                        bracketCount--;
                        break;
                    case TokenType.TableColumn:
                        if (string.IsNullOrEmpty(tableAddress.ColumnName1))
                        {
                            tableAddress.ColumnName1=t.Value;
                        }
                        else
                        {
                            tableAddress.ColumnName2 = t.Value;
                        }
                        break;
                    case TokenType.TablePart:
                        if (string.IsNullOrEmpty(tableAddress.TablePart1))
                        {
                            tableAddress.TablePart1 = t.Value;
                        }
                        else
                        {
                            tableAddress.TablePart2 = t.Value;
                        }
                        break;
                    case TokenType.Colon:
                    case TokenType.Comma:
                        break;
                    default:
                        throw new InvalidFormulaException($"Invalid Table Formula in cell {parsingContext.CurrentCell.Address}");
                }
                //adr += exps[i];
                exps.RemoveAt(i);
                if (bracketCount == 0) break;
            }
            if (extRef <= 0)
            {
                tableAddress.SetTableAddress(parsingContext.Package);
            }
            else
            {
                if(extRef <= parsingContext.Package.Workbook.ExternalLinks.Count)
                {
                    var extWb = parsingContext.Package.Workbook.ExternalLinks[extRef-1].As.ExternalWorkbook;
                    if(extWb != null && extWb.Package!=null)
                    {
                        tableAddress.SetTableAddress(extWb.Package);
                    }
                }
            }
            exps.Insert(i, new Token(tableAddress.WorksheetAddress, TokenType.ExcelAddress));
        }
        private static void ExtractArray(IList<Token> exps, int i, out IRangeInfo range, ParsingContext parsingContext)
        {
            exps.RemoveAt(i);
            var matrix = new List<List<object>>();   
            var array = new List<object>();
            matrix.Add(array);
            var arrayStr= new StringBuilder();
            while (i < exps.Count && exps[i].TokenType != TokenType.ClosingEnumerable)
            {
                var t = exps[i];
                switch (t.TokenType)
                {
                    case TokenType.Boolean:
                        array.Add(bool.Parse(t.Value));
                        break;
                    case TokenType.Integer:
                        array.Add(int.Parse(t.Value));
                        break;
                    case TokenType.Decimal:
                        array.Add(double.Parse(t.Value, NumberStyles.Number, CultureInfo.InvariantCulture));
                        break;
                    case TokenType.StringContent:
                        array.Add(t.Value.Substring(1, t.Value.Length-2).Replace("\"\"","\"")); //Remove double quotes.
                        break;
                    case TokenType.SemiColon:
                        array = new List<object>();
                        matrix.Add(array);
                        break;
                    case TokenType.ClosingEnumerable:
                    case TokenType.Comma:
                        break;
                    case TokenType.NAError:
                        array.Add(ExcelErrorValue.Create(eErrorType.NA));
                        break;
                    case TokenType.InvalidReference:
                        array.Add(ExcelErrorValue.Create(eErrorType.Ref));
                        break;
                    case TokenType.NumericError:
                        array.Add(ExcelErrorValue.Create(eErrorType.Num));
                        break;
                    case TokenType.ValueDataTypeError:
                        array.Add(ExcelErrorValue.Create(eErrorType.Value));
                        break;
                    default:
                        throw new InvalidFormulaException("Array contains invalid tokens. Cell "+ parsingContext.CurrentCell.WorksheetIx);
                }
                arrayStr.Append(t.Value);
                exps.RemoveAt(i);
            }
            if(i==exps.Count || exps[i].TokenType != TokenType.ClosingEnumerable)
            {
                throw new InvalidFormulaException("Array is not closed. Cell " + parsingContext.CurrentCell.WorksheetIx);
            }
            exps.RemoveAt(i);
            exps.Insert(i, new Token(arrayStr.ToString(), TokenType.Array));
            range = new InMemoryRange(matrix);
        }

        private void ExecFunc(Token t, FormulaCell cell)
        {
            var f = _parsingContext.Configuration.FunctionRepository.GetFunction(t.Value);
            var args = GetFunctionArguments(cell);
            var compiler = _functionCompilerFactory.Create(f);
            var result = compiler.Compile(args);
            PushResult(cell, result);
        }

        private void PushResult(FormulaCell cell, CompileResult result)
        {
            switch (result.DataType)
            {
                case DataType.Boolean:
                    cell._expressionStack.Push(new BooleanExpression(result, _parsingContext));
                    break;
                case DataType.Integer:
                    cell._expressionStack.Push(new DecimalExpression(result, _parsingContext));
                    break;
                case DataType.Decimal:
                    cell._expressionStack.Push(new DecimalExpression(result, _parsingContext));
                    break;
                case DataType.String:
                    cell._expressionStack.Push(new StringExpression(result, _parsingContext));
                    break;
                case DataType.ExcelRange:
                    cell._expressionStack.Push(new RangeExpression(result, _parsingContext));
                    break;
            }
        }

        private IList<Expression> GetFunctionArguments(FormulaCell cell)
        {
            var list = new List<Expression>();
            var pos = cell._funcStackPosition.Pop();
            var s = cell._expressionStack;
            while (s.Count > pos._startPos)
            {
                var si = s.Pop();
                si.Status |= ExpressionStatus.FunctionArgument;
                list.Insert(0, si);
            }
            return list;
        }
        private void AddExpressionOrApplyOperator(IList<Expression> precompiledExps, Token opToken, FormulaCell cell)
        {
            var v1 = cell._expressionStack.Pop();
            var v2 = cell._expressionStack.Pop();

            if (OperatorsDict.Instance.TryGetValue(opToken.Value, out IOperator op))
            {
                if ((v1.Status == ExpressionStatus.CanCompile && 
                    v2.Status == ExpressionStatus.CanCompile) ||
                    (v1.Status == ExpressionStatus.IsAddress &&
                    v2.Status == ExpressionStatus.IsAddress && op.Operator==Operators.Colon))
                {
                    var c1 = v1.Compile();
                    var c2 = v2.Compile();

                    var result = op.Apply(c2, c1, _parsingContext);
                    PushResult(cell, result);
                }
                else
                {
                    if (v1.Status == ExpressionStatus.OnExpressionList || v2.Status == ExpressionStatus.OnExpressionList)
                    {
                        precompiledExps[precompiledExps.Count - 1].Operator = op.Operator;
                    }
                    else
                    {
                        v2.Operator = op.Operator;
                    }
                    
                    if (v2.Status != ExpressionStatus.OnExpressionList)
                    {
                        v2.Status = ExpressionStatus.OnExpressionList;
                        precompiledExps.Add(v2);
                    }
                    if (v1.Status != ExpressionStatus.OnExpressionList)
                    {
                        precompiledExps.Add(v1);
                        v1.Status = ExpressionStatus.OnExpressionList;
                    }
                    cell._expressionStack.Push(v1);
                }
            }
            else
            {
                throw new InvalidFormulaException($"Invalid operator {opToken.Value}");
            }
        }

        private void ApplyOperator(Token opToken, FormulaCell cell)
        {
            var v1 = cell._expressionStack.Pop();
            var v2 = cell._expressionStack.Pop();
            
            var c1 = v1.Compile();
            var c2 = v2.Compile();

            if (OperatorsDict.Instance.TryGetValue(opToken.Value, out IOperator op))
            {
                var result = op.Apply(c2, c1, _parsingContext);
                PushResult(cell, result);
            }
        }
    }
}
