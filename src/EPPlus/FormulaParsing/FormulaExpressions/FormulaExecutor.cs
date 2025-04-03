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
using OfficeOpenXml.Table;
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

        internal FormulaExecutor(ParsingContext parsingContext)
        {
            _parsingContext = parsingContext;
        }

        internal static RpnTokens CreateRPNTokens(IList<Token> tokens)
        {
            var bracketCount = 0;
            var lastOpeningParanthesesPos = -1;
            var lastCommaPos=-1;
            var operators = OperatorsDict.Instance;
            Stack<Token> operatorStack = new Stack<Token>();
            Stack<int> lambdas = new Stack<int>();
            var rpnTokens = new List<Token>();
            var hasLambda = false;
            for (int i = 0; i < tokens.Count; i++)
            {
                Token token = tokens[i];
                switch (token.TokenType)
                {
                    case TokenType.OpeningParenthesis:
                        lastOpeningParanthesesPos = rpnTokens.Count;
                        operatorStack.Push(token);
                        break;
                    case TokenType.ClosingParenthesis:
                        if (operatorStack.Count > 0)
                        {                            
                            var o = operatorStack.Pop();
                            var noOperator = true;
                            while (o.TokenType != TokenType.OpeningParenthesis)
                            {
                                rpnTokens.Add(o);
                                if (operatorStack.Count == 0) throw new InvalidOperationException("No closing parenthesis");
                                o = operatorStack.Pop();
                                noOperator = false;
                            }

                            if (operatorStack.Count > 0 && operatorStack.Peek().TokenType == TokenType.Function)
                            {
                                rpnTokens.Add(operatorStack.Pop());
                                noOperator = false;
                            }

                            //if (noOperator && lastOpeningParanthesesPos > -1 && lastCommaPos > -1 && rpnTokens.Count - lastOpeningParanthesesPos > 1) // Handle comma separated addresses to become one expression. For Example SUM( (A1,A3) ,A5) where A1,A3 should become one argument to the SUM function.
                            //{
                            //    var sb = new StringBuilder();
                            //    while(rpnTokens.Count > lastOpeningParanthesesPos)
                            //    {
                            //        sb.Append(rpnTokens[lastOpeningParanthesesPos].Value);
                            //        rpnTokens.RemoveAt(lastOpeningParanthesesPos);
                            //    }
                            //    rpnTokens.Add(new Token(sb.ToString(), TokenType.ExcelAddress));

                            //}
                            lastOpeningParanthesesPos = -1;
                            lastCommaPos = -1;
                        }
                        break;
                    case TokenType.Operator:
                    case TokenType.Negator:
                        if(token.TokenType == TokenType.Operator && i > 0 && i < tokens.Count-2 && token.Value==":" && tokens[i-1].Value=="]" && tokens[i+1].Value=="[")
                        {
                            rpnTokens.Add(token);
                            break;
                        }
                        if (operatorStack.Count > 0)

                        {
                            var o2 = operatorStack.Peek();
                            while ((o2.TokenType == TokenType.Operator && token.TokenType!=TokenType.Negator &&
                                operators[o2.Value].Precedence <= operators[token.Value].Precedence) 
                                || 
                                (o2.TokenType == TokenType.Negator && 
                                token.TokenType != TokenType.Negator && 
                                operators[token.Value].Precedence > Operator.PrecedenceColon))
                            {
                                rpnTokens.Add(operatorStack.Pop());
                                if (operatorStack.Count == 0) break;
                                o2 = operatorStack.Peek();
                            }
                        }
                        operatorStack.Push(token);
                        break;

                    case TokenType.Function:
                        rpnTokens.Add(new Token(token.Value,TokenType.StartFunctionArguments));
                        operatorStack.Push(token);
                        break;
                    case TokenType.Comma:
                    case TokenType.CommaLambda:
                        if(operatorStack.Count > 0 && bracketCount == 0) //If inside a table 
                        {
                            var op = operatorStack.Peek().TokenType;
                            while (op == TokenType.Operator || op == TokenType.Negator)
                            {
                                rpnTokens.Add(operatorStack.Pop());
                                if(operatorStack.Count == 0) break;
                                op = operatorStack.Peek().TokenType;
                            }
                        }
                        if (token.TokenType == TokenType.CommaLambda)
                        {
                            hasLambda = true;
                        }
                        lastCommaPos = i;
                        rpnTokens.Add(token);
                        break;
                    case TokenType.OpeningBracket:
                        bracketCount++;
                        rpnTokens.Add(token);
                        break;
                    case TokenType.ClosingBracket:
                        bracketCount--;
                        rpnTokens.Add(token);
                        break;
                    default:
                        rpnTokens.Add(token);
                        break;
                }

            }

            while (operatorStack.Count > 0)
            {
                rpnTokens.Add(operatorStack.Pop());
            }
            var result = new RpnTokens
            {
                Tokens = rpnTokens
            };
            if (hasLambda)
            {
                ProcessLambda(result);
            }
            return result;
        }

        private static void ProcessLambda(RpnTokens rpnTokens)
        {
            var lambdaRefs = new Dictionary<int, int>();
            Stack<int> lStack = new Stack<int>();
            for (var i = 0; i < rpnTokens.Count; i++)
            {
                var token = rpnTokens[i];
                if (token.IsLambdaFunction())
                {
                    if (token.TokenType == TokenType.StartFunctionArguments)
                    {
                        lStack.Push(i);
                    }
                    else
                    {
                        lambdaRefs[lStack.Pop()] = i + 1;
                    }
                }
            }
            rpnTokens.LambdaRefs = lambdaRefs;

            // TODO: create a new list and loop through the existing tokens
            // move tokens from the Lambda invoke parenthesis to after the corresponding
            // variable and use the new Assign operator... /MA
            //var TokenList = new List<Token>();
        }

        public static Dictionary<int, Expression> CompileExpressions(ref RpnTokens rpnTokens, ParsingContext parsingContext)
        {
            return CompileExpressions(null, ref rpnTokens, parsingContext);
        }


        public static Dictionary<int, Expression> CompileExpressions(RpnFormula rpnFormula, ref RpnTokens rpnTokens, ParsingContext parsingContext)
        {
            short extRefIx = short.MinValue;
            int wsIx = int.MinValue;
            var stack = new Stack<FunctionExpression>();
            var expressions = new Dictionary<int, Expression>();
            var isInLambdaCalculation = false;
            LambdaTokensExpression lambdaCalculationExpression = null;
            var tokens = rpnTokens.Tokens;
            var variableStorage = rpnFormula != null ? rpnFormula.VariableStorage : new VariableStorage.VariableStorageManager();
            var lambdaLevel = 0;
            for (int tokenIx = 0; tokenIx < tokens.Count; tokenIx++)
            {
                var t = tokens[tokenIx];
                if (isInLambdaCalculation)
                {
                    var isLambdaToken = t.IsLambdaFunction();
                    if(isLambdaToken && t.TokenType == TokenType.StartFunctionArguments)
                    {
                        lambdaLevel++;
                        if (rpnFormula != null)
                        {
                            rpnFormula.LambdaSettings.AddLambdaToken(tokenIx);
                        }
                        lambdaCalculationExpression.AddLambdaToken(t);
                        continue;
                    }
                    else if (isLambdaToken && t.TokenType == TokenType.Function && lambdaLevel > 0)
                    {
                        if(lambdaLevel > 0)
                        {
                            if (rpnFormula != null)
                            {
                                rpnFormula.LambdaSettings.AddLambdaToken(tokenIx);
                            }
                            lambdaCalculationExpression.AddLambdaToken(t);
                            lambdaLevel--;
                            continue;
                        }
                    }
                    else if (!(t.TokenType == TokenType.Function && isLambdaToken) || lambdaLevel > 0)
                    {
                        if (rpnFormula != null)
                        {
                            rpnFormula.LambdaSettings.AddLambdaToken(tokenIx);
                        }
                        lambdaCalculationExpression.AddLambdaToken(t);
                        continue;
                    }  
                }
                if (rpnTokens.HasLambdaRefs && rpnTokens.LambdaRefs.ContainsKey(tokenIx))
                {
                    var tknIx = rpnTokens.LambdaRefs[tokenIx] + 1;
                    var list = new List<Token>();
                    tknIx = tknIx <= rpnTokens.Count() - 1 ? tknIx : rpnTokens.Count() - 1;
                    var sTkn = rpnTokens[tknIx - 1];
                    if (sTkn.TokenType == TokenType.LambdaInvokeArgsStart)
                    {
                        var tkn = rpnTokens[tknIx];
                        while (tkn.TokenType != TokenType.LambdaInvokeArgsEnd)
                        {
                            list.Add(tkn);
                            tkn = rpnTokens[++tknIx];
                        }
                    }

                }
                switch (t.TokenType)
                {
                    case TokenType.Boolean:
                        expressions.Add(tokenIx, new BooleanExpression(t.Value, parsingContext));
                        break;
                    case TokenType.Integer:
                        expressions.Add(tokenIx, new IntegerExpression(t.Value, parsingContext));
                        break;
                    case TokenType.Decimal:
                        expressions.Add(tokenIx, new DecimalExpression(t.Value, parsingContext));
                        break;
                    case TokenType.StringContent:
                        expressions.Add(tokenIx, new StringExpression(t.Value, parsingContext));
                        break;
                    case TokenType.CellAddress:
                    case TokenType.FullColumnAddress:
                    case TokenType.FullRowAddress:
                        if (tokenIx > 1 && tokens[tokenIx - 1].TokenTypeIsAddress && tokens[tokenIx + 1].Value == ":" && tokens[tokenIx + 1].TokenType == TokenType.Operator)
                        {
                            //We have a two cell addresses with with a colon. Remove tokens and replace with full column address, for example A1:C2.
                            var e = expressions[tokenIx - 1];
                            e.MergeAddress(t.Value);
                            tokens.RemoveAt(tokenIx - 1);
                            tokens.RemoveAt(tokenIx);
                            tokenIx--;
                            tokens[tokenIx] = new Token(e.GetAddress()[0].WorksheetAddress, TokenType.ExcelAddress);
                        }
                        else
                        {
                            expressions.Add(tokenIx, new RangeExpression(t.Value, parsingContext, extRefIx, wsIx));
                        }
                        extRefIx = short.MinValue;
                        wsIx = int.MinValue;
                        break;
                    case TokenType.ExcelAddress:
                        var a = new ExcelAddressBase(t.Value);
                        if(a.Addresses?.Count>1)
                        {
                            expressions.Add(tokenIx, new MultiRangeExpression(a, parsingContext));
                        }
                        else
                        {
                            expressions.Add(tokenIx, new RangeExpression(a.AsFormulaRangeAddress(parsingContext)));
                        }
                        extRefIx = short.MinValue;
                        wsIx = int.MinValue;
                        break;
                    case TokenType.NameValue:                        
                        expressions.Add(tokenIx, new NamedValueExpression(t.Value, parsingContext, extRefIx, wsIx));
                        wsIx = int.MinValue;
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
                        ExtractTableAddress(extRefIx, wsIx, tokens, tokenIx, out FormulaTableAddress tableAddress, parsingContext);                        
                        expressions.Add(tokenIx, new TableAddressExpression(tableAddress, parsingContext));
                        break;
                    case TokenType.OpeningEnumerable:
                        ExtractArray(tokens, tokenIx, out IRangeInfo rangInfo, parsingContext);
                        expressions.Add(tokenIx, new EnumerableExpression(rangInfo, parsingContext));
                        break;
                    case TokenType.ParameterVariableDeclaration:
                        var variableFunction = stack.Peek() as VariableFunctionExpression;
                        expressions.Add(tokenIx, new VariableExpression(t.Value, variableFunction, true));
                        break;
                    case TokenType.ParameterVariable:
                        foreach(var exp in stack)
                        {
                            if(exp is VariableFunctionExpression vfeExp)
                            {
                                if(vfeExp.VariableIsDeclared(t.Value) && !expressions.ContainsKey(tokenIx))
                                {
                                    expressions.Add(tokenIx, new VariableExpression(t.Value, vfeExp, false));
                                }
                            }
                        }
                        break;
                    case TokenType.StartFunctionArguments:
                        FunctionExpression func = default;
                        if(t.IsLetFunction())
                        {
                            func = new LetFunctionExpression(t.Value, variableStorage, parsingContext, tokenIx);
                        }
                        else if(t.IsLambdaFunction())
                        {
                            func = new LambdaFunctionExpression(t.Value, variableStorage, parsingContext, tokenIx);
                        }
                        else if(t.IsBuiltInFunction(parsingContext.Configuration.FunctionRepository))
                        {
                            func = new FunctionExpression(t.Value, parsingContext, tokenIx);
                        }
                        else if (parsingContext.Package.Workbook.Names.ContainsKey(t.Value) || parsingContext.CurrentWorksheet.Names.ContainsKey(t.Value))
                        {
                            var name = parsingContext.Package.Workbook.Names[t.Value];
                            if (name == null)
                            {
                                name = parsingContext.CurrentWorksheet.Names[t.Value];
                            }
                            if (name != null)
                            {
                                var lambdaFormulaCandidate = name.Formula;
                                if (!string.IsNullOrEmpty(lambdaFormulaCandidate) && lambdaFormulaCandidate.ToLower().StartsWith("lambda") || lambdaFormulaCandidate.ToLower().StartsWith("_xlfn.lambda"))
                                {
                                    func = new LambdaNameFunctionExpression(t.Value, lambdaFormulaCandidate, variableStorage, parsingContext, tokenIx);
                                }
                            }
                        }
                        if (func == null) func = new FunctionExpression(t.Value, parsingContext, tokenIx);             
                        expressions.Add(tokenIx, func);
                        if(tokenIx <= tokens.Count && tokens[tokenIx + 1].TokenType != TokenType.Function) // Check that the function has any argument
                        {
                            func.AddArgument(tokenIx);
                        }
                        stack.Push(func);
                        break;
                    case TokenType.Comma:
                        if (stack.Count > 0)
                        {
                            stack.Peek().AddArgument(tokenIx);
                        }
                        break;
                    case TokenType.CommaLambda:
                        isInLambdaCalculation = true;
                        lambdaCalculationExpression = new LambdaTokensExpression(parsingContext, variableStorage.AddNewScope());
                        expressions.Add(tokenIx, lambdaCalculationExpression);
                        if (stack.Count > 0)
                        {
                            stack.Peek().AddArgument(tokenIx);
                        }
                        break;
                    case TokenType.Function:
                        var f = stack.Pop();
                        f._endPos= tokenIx;
                        if (f.IsLambda && isInLambdaCalculation)
                        {
                            lambdaCalculationExpression = null;
                            isInLambdaCalculation = false;
                        }
                        break;
                    case TokenType.InvalidReference:
                        expressions.Add(tokenIx, ErrorExpression.RefError);
                        wsIx = int.MinValue;
                        break;
                }
            }
            return expressions;
        }

        private static void ExtractTableAddress(int extRef, int wsIx, IList<Token> exps, int i, out FormulaTableAddress tableAddress, ParsingContext parsingContext)
        {
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
    }
}
