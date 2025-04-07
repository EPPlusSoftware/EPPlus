/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/12/2024         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaCalculator
    {
        public LambdaCalculator(List<Token> lambdaTokens, VariableStorageScope scope)
        {
            _originalTokens = lambdaTokens;
            _scope = scope;
            for(var i = 0; i < _originalTokens.Count; i++)
            {
                var t = _originalTokens[i];
                if(t.TokenType == TokenType.ParameterVariable)
                {
                    _variableIndexes.Add(i);
                }
            }
        }

        private List<int> _variableIndexes = new List<int>();
        private List<CompileResult> _variables;
        private VariableStorageScope _scope;
        private readonly List<Token> _originalTokens;
        private List<Token> _currentTokens;
        private int _nVariablesSet = 0;

        public short NumberOfVariables => _variables != null ? Convert.ToInt16(_variables.Count()) : (short)0;

        public bool IsReadyForCalc => _nVariablesSet >= NumberOfVariables;

        public void BeginCalculation()
        {
            CloneTokens();
        }

        public void ResetVariables()
        {
            _nVariablesSet = 0;
        }

        internal string GetDebugInfo()
        {
            if(_variables == null)
            {
                return "#variables: 0";
            }
            var sb = new StringBuilder();
            sb.Append("#variables: " + _variables.Count() + ", ");
            foreach (var v in _variables)
            {
                sb.Append(v.Result);
            }
            return sb.ToString();
        }

        public void SetVariables(List<CompileResult> variables)
        {
            _variables = variables;
        }

        public void SetVariableValue(int index, object value, DataType dt, ParsingContext ctx)
        {
            var variable = _variables[index];
            foreach(var ix in _variableIndexes)
            {
                var t = _currentTokens[ix];
                if (string.Compare(t.Value, variable.Result.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    if(value is LambdaCalculator lc)
                    {
                        var cr = lc.Execute(ctx);
                        value = cr.Result;
                        dt = cr.DataType;
                    }
                    var tt = DataTypeToTokenType(dt, value);
                    if (tt == TokenType.Unrecognized) tt = TokenType.StringContent;
                    var tokenValue = Convert.ToString(value, CultureInfo.CurrentCulture);
                    if(tt == TokenType.StringContent)
                    {
                        tokenValue = $"\"{tokenValue}\"";
                    }
                    _currentTokens[ix] = new Token(tokenValue, tt);
                }
            }
            _nVariablesSet++;
        }

        public CompileResult Execute(ParsingContext ctx)
        {
            var formula = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column);
            var rpnTokens = new RpnTokens { Tokens = _currentTokens };
            formula.SetTokens(rpnTokens, ctx);
            var chain = new RpnOptimizedDependencyChain(ctx.CurrentWorksheet.Workbook, ctx.CalcOption);
            var result = RpnFormulaExecution.ExecutePartialFormula(chain, formula, ctx.CalcOption, false);
            return CompileResultFactory.Create(result);
        }

        private void CloneTokens()
        {
            _currentTokens = new List<Token>();
            foreach(var token in _originalTokens)
            {
                _currentTokens.Add(new Token(token.Value, token.TokenType));
            }
        }

        private TokenType DataTypeToTokenType(DataType dt, object obj)
        {
            switch (dt)
            {
                case DataType.Boolean:
                    return TokenType.Boolean;
                case DataType.Date:
                case DataType.Integer:
                    return TokenType.Integer;
                case DataType.Decimal:
                    return TokenType.Decimal;
                case DataType.String:
                    return TokenType.StringContent;
                case DataType.Time:
                    return TokenType.Decimal;
                case DataType.ExcelRange:
                    return TokenType.ExcelAddress;
                case DataType.ExcelError:
                    switch(obj.ToString().ToUpper())
                    {
                        case "#N/A":
                            return TokenType.NAError;
                        case "#NAME!":
                            return TokenType.NameError;
                        case "#NUM!":
                            return TokenType.NumericError;
                        default:
                            return TokenType.ValueDataTypeError;
                    }
                default:
                    return TokenType.Unrecognized;
            }
        }
    }
}
