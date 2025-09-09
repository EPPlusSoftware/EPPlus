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
    /// <summary>
    /// Stores the Lambda tokens and can be called multiple times with different parameters. 
    /// For each time the tokens are converted RPN tokens and runs through the calculation 
    /// via the new <see cref="RpnFormulaExecution.ExecutePartialFormula(RpnOptimizedDependencyChain, RpnFormula, ExcelCalculationOption, bool, VariableStorageScope)"/> method.
    /// </summary>
    internal class LambdaCalculator
    {
        /// <summary>
        /// Tokens representing the LAMBDA expression. For example if the LAMBDA function
        /// is LAMBDA(x, 1, x + 1) the tokens supplied to the LambdaCalculator will be x + 1
        /// </summary>
        /// <param name="lambdaTokens"></param>
        /// <param name="scope"></param>
        /// <param name="formula"></param>
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

        /// <summary>
        /// Resets the calculator for a new calculation
        /// </summary>
        public void BeginCalculation()
        {
            CloneTokens();
        }

        /// <summary>
        /// Resets the counting of variables
        /// </summary>
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

        /// <summary>
        /// Sets the variables values, creates new variables if not present.
        /// </summary>
        /// <param name="variables">List of variables</param>
        /// <param name="ctx">The parsing context</param>
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
        }

        /// <summary>
        /// Sets a variable's value in the variable scope.
        /// </summary>
        /// <param name="index">0-based index</param>
        /// <param name="value">name of the variable</param>
        /// <param name="dt">DataType of the variable</param>
        /// <param name="ctx">The parsing context</param>
        public void SetVariableValue(int index, object value, DataType dt, ParsingContext ctx)
        {
            SetVariableValue(index, value, dt, ctx, null);
        }

        /// <summary>
        /// Sets a variable's value in the variable scope.
        /// </summary>
        /// <param name="index">0-based index</param>
        /// <param name="value">name of the variable</param>
        /// <param name="dt">DataType of the variable</param>
        /// <param name="ctx">The parsing context</param>
        /// <param name="address">The cell address if applicable</param>
        public void SetVariableValue(int index, object value, DataType dt, ParsingContext ctx, FormulaRangeAddress address)
        {
            var variable = _variables[index];
            var variableName = variable.VariableName;
            if(value is IRangeInfo ir && !ir.IsMulti)
            {
                var v = ir.GetOffset(0, 0);
                var tmpCr = CompileResultFactory.Create(v);
                dt = tmpCr.DataType;
                value = v;
            }
            var cr = address != null ? new AddressCompileResult(value, dt, address) : new CompileResult(value, dt);
            _variables[index] = new VariableCompileResult(variableName, value, dt, address);
            _scope.SetVariableValue(variableName, cr);
            _nVariablesSet++;
        }

        /// <summary>
        /// Executes the Lambda function and returns the result
        /// </summary>
        /// <param name="ctx"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public CompileResult Execute(ParsingContext ctx)
        {
            if(_currentTokens == null)
            {
                throw new InvalidOperationException("LambdaCalculator.Execute was called without having initialized it with BeginCalculation. OriginalTokens.Count = " + _originalTokens.Count + ", # variables = " + (_variables != null ? _variables.Count : 0));
            }
            if(_nVariablesSet < NumberOfVariables)
            {
                var result = ExcelErrorValue.Create(eErrorType.Value);
                return CompileResultFactory.CreateDynamicArrayResult(result);
            }
            var formula = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column, ctx.VariableStorage);
            formula.IgnoreCaching = true;
            formula.ExpressionStack = _formula.ExpressionStack;
            formula.FunctionStack = _formula.FunctionStack;
            var rpnTokens = new RpnTokens { Tokens = _currentTokens, Scope = _scope };
            formula.SetTokens(rpnTokens, ctx, _scope);
            // SetTokens clears the variable storage...
            ctx.VariableStorage.Push(_scope);
            var chain = ctx.DependencyChain;
            var compileResult = RpnFormulaExecution.ExecutePartialFormula(chain, formula, ctx.CalcOption, false);
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
