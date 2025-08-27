/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/20/2025         EPPlus Software AB       Initial release EPPlus 8.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaCalculator
    {
        public LambdaCalculator(List<Token> lambdaTokens, VariableStorageScope scope, RpnFormula formula)
        {
            _originalTokens = lambdaTokens;
            _scope = scope;
            _formula = formula;
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
        private List<VariableCompileResult> _variables = new List<VariableCompileResult>();
        private VariableStorageScope _scope;
        private readonly List<Token> _originalTokens;
        private List<Token> _currentTokens;
        private int _nVariablesSet = 0;
        private readonly RpnFormula _formula;

        public short NumberOfVariables => _variables != null ? Convert.ToInt16(_variables.Count()) : (short)0;

        public bool IsReadyForCalc
        {
            get
            {
                foreach(var variable in _variables)
                {
                    if (_scope.ContainsVariable(variable.VariableName)) continue;
                    return false;
                }
                return true;
            }    
        }

        public bool IsCompileLambdaName
        {
            get; set;
        }

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

        public List<Token> GetCurrentTokens()
        {
            return _currentTokens;
        }

        public void SetVariables(List<VariableCompileResult> variables, ParsingContext ctx)
        {
            if (variables == null || !variables.Any()) return;
            for(var i = 0; i < variables.Count; i++)
            {
                if (!_variables.Any(x => x.VariableName == variables[i].VariableName))
                {
                    _variables.Add(variables[i]);
                }
                if (variables[i].DataType != DataType.LambdaVariableDeclaration && variables[i].DataType != DataType.Variable)
                {
                    SetVariableValue(i, variables[i].ResultValue, variables[i].DataType, ctx);
                }
            }
            #region Old code
            //for(var i = 0; i < _variables.Count; i++)
            //{
            //    if(i < variables.Count && variables[i].DataType == DataType.Variable)
            //    {
            //        var name = _variables[i].GetResultValue();
            //        var variableValue = variables[i].ResultValue;
            //        var dt = variables[i].DataType;
            //        _variables[i] = new CompileResult(name, dt);
            //        _scope.SetVariableValue(name, new CompileResult(variableValue, dt));
            //    }
            //}
            #endregion
        }

        public void SetVariableValue(int index, object value, DataType dt, ParsingContext ctx)
        {
            SetVariableValue(index, value, dt, ctx, null);
        }

        public void SetVariableValue(int index, object value, DataType dt, ParsingContext ctx, FormulaRangeAddress address)
        {
            var variable = _variables[index];
            var variableName = variable.VariableName;
            var cr = address != null ? new AddressCompileResult(value, dt, address) : new CompileResult(value, dt);
            _variables[index] = new VariableCompileResult(variableName, value, dt, address);
            _scope.SetVariableValue(variableName, cr);
            #region Old code
            //var val = address != null ? address.Address : value;
            //dt = address != null ? DataType.ExcelRange : dt;
            //if(dt == DataType.ExcelRange && val is string adr)
            //{
            //    var fAdr = new FormulaRangeAddress(ctx, adr);
            //    val = ctx.ExcelDataProvider.GetRange(fAdr);
            //}
            //var compileResult = new CompileResult(val, dt);
            //_scope.SetVariableValue(variableName, compileResult);
            //foreach(var ix in _variableIndexes)
            //{
            //    var t = _currentTokens[ix];
            //    if (string.Compare(t.Value, variableName, StringComparison.OrdinalIgnoreCase) == 0)
            //    {
            //        if(value is LambdaCalculator lc)
            //        {
            //            var cr = lc.Execute(ctx);
            //            value = cr.Result;
            //            dt = cr.DataType;
            //        }
            //        var tt = DataTypeToTokenType(dt, val);
            //        if (tt == TokenType.Unrecognized) tt = TokenType.StringContent;
            //        if(val is IRangeInfo rng)
            //        {
            //            if(rng.IsInMemoryRange)
            //            {
            //                var imRng = rng as InMemoryRange;
            //                var tokens = imRng.SerializeToTokens();
            //                var preceedingTokens = _currentTokens.Take(ix);
            //                var trailingTokens = _currentTokens.Skip(ix + 1);
            //                _currentTokens = new List<Token>(preceedingTokens);
            //                _currentTokens.AddRange(tokens);
            //                _currentTokens.AddRange(trailingTokens);
            //                return;
            //            }
            //            else
            //            {
            //                value = rng.Address.Address;
            //            }
            //        }
            //        var tokenValue = Convert.ToString(val, CultureInfo.CurrentCulture);
            //        if(tt == TokenType.StringContent)
            //        {
            //            tokenValue = $"\"{tokenValue}\"";
            //        }
            //        _currentTokens[ix] = new Token(tokenValue, tt);
            //    }
            //}
            #endregion
            _nVariablesSet++;
        }

        public CompileResult Execute(ParsingContext ctx)
        {
            if(_currentTokens == null)
            {
                throw new InvalidOperationException("LambdaCalculator.Execute was called without having initialized it with BeginCalculation. OriginalTokens.Count = " + _originalTokens.Count + ", # variables = " + (_variables != null ? _variables.Count : 0));
            }
            var formula = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column, ctx.VariableStorage);
            formula.IgnoreCaching = true;
            formula.ExpressionStack = _formula.ExpressionStack;
            formula.FunctionStack = _formula.FunctionStack;
            var rpnTokens = new RpnTokens { Tokens = _currentTokens, Scope = _scope };
            formula.SetTokens(rpnTokens, ctx);
            var chain = ctx.DependencyChain;
            var compileResult = RpnFormulaExecution.ExecutePartialFormula(chain, formula, ctx.CalcOption, false);
            if (_formula.ExpressionStack.Count > 0 && _formula.ExpressionStack.Peek() is LambdaCalculationExpression lce)
            {
                _formula.ExpressionStack.Pop();
            }
            return CompileResultFactory.CreateDynamicArrayResult(compileResult.Result, compileResult.Address, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
        }

        private void CloneTokens()
        {
            _currentTokens = new List<Token>();
            foreach(var token in _originalTokens)
            {
                _currentTokens.Add(new Token(token.Value, token.TokenType));
            }
        }
    }
}
