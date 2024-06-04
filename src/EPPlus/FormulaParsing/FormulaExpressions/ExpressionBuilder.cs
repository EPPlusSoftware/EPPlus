﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2024         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal static class ExpressionBuilder
    {
        public static Dictionary<int, Expression> BuildExpressions(ref IList<Token> tokens, ParsingContext parsingContext)
        {
            return BuildExpressions(null, ref tokens, parsingContext);
        }

        public static Dictionary<int, Expression> BuildExpressions(RpnFormula rpnFormula, ref IList<Token> tokens, ParsingContext parsingContext)
        {
            short extRefIx = short.MinValue;
            int wsIx = int.MinValue;
            var stack = new Stack<FunctionExpression>();
            var expressions = new Dictionary<int, Expression>();
            var isInLambdaCalculation = false;
            LambdaTokensExpression lambdaCalculationExpression = null;
            for (int tokenIx = 0; tokenIx < tokens.Count; tokenIx++)
            {
                var t = tokens[tokenIx];
                if(isInLambdaCalculation && t.TokenType != TokenType.Function)
                {
                    if(rpnFormula != null)
                    {
                        rpnFormula.AddLambdaToken(tokenIx);
                    }
                    lambdaCalculationExpression.AddLambdaToken(t);
                    continue;
                }
                var i = tokenIx;
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
                            tokenIx--;
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
                        ExtractTableAddress(extRefIx, wsIx, tokens, i, out FormulaTableAddress tableAddress, parsingContext);                        
                        expressions.Add(i, new TableAddressExpression(tableAddress, parsingContext));
                        break;
                    case TokenType.OpeningEnumerable:
                        ExtractArray(tokens, i , out IRangeInfo rangInfo, parsingContext);
                        expressions.Add(i, new EnumerableExpression(rangInfo, parsingContext));
                        break;
                    case TokenType.ParameterVariableDeclaration:
                        var variableFunction = stack.Peek() as VariableFunctionExpression;
                        expressions.Add(i, new VariableExpression(t.Value, variableFunction, true));
                        break;
                    case TokenType.ParameterVariable:
                        foreach(var exp in stack)
                        {
                            if(exp is VariableFunctionExpression vfeExp)
                            {
                                if(vfeExp.VariableIsDeclared(t.Value))
                                {
                                    expressions.Add(i, new VariableExpression(t.Value, vfeExp, false));
                                }
                            }
                        }
                        break;
                    case TokenType.StartFunctionArguments:
                        var isLet = !string.IsNullOrEmpty(t.Value) && (string.Compare(t.Value, "_xlfn.LET", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(t.Value, "LET", StringComparison.OrdinalIgnoreCase) == 0);
                        var isLambda = !string.IsNullOrEmpty(t.Value) && (string.Compare(t.Value, "_xlfn.LAMBDA", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(t.Value, "LAMBDA", StringComparison.OrdinalIgnoreCase) == 0);
                        var func = isLet ? 
                            new LetFunctionExpression(t.Value, stack, parsingContext, i) :
                                isLambda ?
                                    new LambdaFunctionExpression(t.Value, stack, parsingContext, i) :
                                    new FunctionExpression(t.Value, parsingContext, i);
                        expressions.Add(i, func);
                        if(i <= tokens.Count && tokens[i+1].TokenType != TokenType.Function) // Check that the function has any argument
                        {
                            func.AddArgument(i);
                        }
                        stack.Push(func);
                        break;
                    case TokenType.Comma:
                        if (stack.Count > 0)
                        {
                            stack.Peek().AddArgument(i);
                        }
                        break;
                    case TokenType.CommaLambda:
                        isInLambdaCalculation = true;
                        lambdaCalculationExpression = new LambdaTokensExpression(parsingContext);
                        expressions.Add(i, lambdaCalculationExpression);
                        if (stack.Count > 0)
                        {
                            stack.Peek().AddArgument(i);
                        }
                        break;
                    case TokenType.Function:
                        var f = stack.Pop();
                        f._endPos= i;
                        if(f.IsLambda && isInLambdaCalculation)
                        {
                            lambdaCalculationExpression = null;
                            isInLambdaCalculation = false;
                        }
                        break;
                    case TokenType.InvalidReference:
                        expressions.Add(i, ErrorExpression.RefError);
                        break;
                }
            }
            return expressions;
        }

        private static void ExtractTableAddress(int extRef, int wsIx, IList<Token> exps, int i, out FormulaTableAddress tableAddress, ParsingContext parsingContext)
        {
            //var adr = exps[i].Value;
            tableAddress = new FormulaTableAddress(parsingContext) {ExternalReferenceIx = extRef, WorksheetIx=wsIx, TableName = exps[i].Value };
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
                            tableAddress.ColumnName1 = ExcelTableColumn.DecodeTableColumnName(t.Value);
                        }
                        else
                        {
                            tableAddress.ColumnName2 = ExcelTableColumn.DecodeTableColumnName(t.Value);
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
                    case TokenType.Comma:
                        break;
                    default:
                        if (t.TokenType == TokenType.Operator && t.Value == ":") break;
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

    //    private void PushResult(FormulaCell cell, CompileResult result)
    //    {
    //        switch (result.DataType)
    //        {
    //            case DataType.Boolean:
    //                cell._expressionStack.Push(new BooleanExpression(result, _parsingContext));
    //                break;
    //            case DataType.Integer:
    //                cell._expressionStack.Push(new DecimalExpression(result, _parsingContext));
    //                break;
    //            case DataType.Decimal:
    //                cell._expressionStack.Push(new DecimalExpression(result, _parsingContext));
    //                break;
    //            case DataType.String:
    //                cell._expressionStack.Push(new StringExpression(result, _parsingContext));
    //                break;
    //            case DataType.ExcelRange:
    //                cell._expressionStack.Push(new RangeExpression(result, _parsingContext));
    //                break;
    //        }
    //    }

    //    private IList<Expression> GetFunctionArguments(FormulaCell cell)
    //    {
    //        var list = new List<Expression>();
    //        var pos = cell._funcStackPosition.Pop();
    //        var s = cell._expressionStack;
    //        while (s.Count > pos._startPos)
    //        {
    //            var si = s.Pop();
    //            si.Status |= ExpressionStatus.FunctionArgument;
    //            list.Insert(0, si);
    //        }
    //        return list;
    //    }
    }
}
