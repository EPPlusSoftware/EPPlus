using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaCalculator
    {
        public LambdaCalculator(List<Token> lambdaTokens)
        {
            _originalTokens = lambdaTokens;
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
        private readonly List<Token> _originalTokens;
        private List<Token> _currentTokens;

        public void BeginCalculation()
        {
            CloneTokens();
        }

        public void SetVariables(List<CompileResult> variables)
        {
            _variables = variables;
        }

        public void SetVariableValue(int index, object value, DataType dt)
        {
            var variable = _variables[index];
            foreach(var ix in _variableIndexes)
            {
                var t = _currentTokens[ix];
                if (string.Compare(t.Value, variable.Result.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    var tt = DataTypeToTokenType(dt, value);
                    var tokenValue = Convert.ToString(value, CultureInfo.CurrentCulture);
                    _currentTokens[ix] = new Token(value.ToString(), tt);
                }
            }
        }

        public CompileResult Execute(ParsingContext ctx)
        {
            var formula = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column);
            formula.SetTokens(_currentTokens, ctx);
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
                    return TokenType.String;
                case DataType.Time:
                    return TokenType.Decimal;
                case DataType.ExcelError:
                    switch(obj.ToString().ToUpper())
                    {
                        case "#NA!":
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
