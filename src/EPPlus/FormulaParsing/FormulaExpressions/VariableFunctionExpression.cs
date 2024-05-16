using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class VariableFunctionExpression : FunctionExpression
    {
        internal VariableFunctionExpression(string tokenValue, ParsingContext ctx, int pos) : base(tokenValue, ctx, pos)
        {
        }

        private readonly Dictionary<string, CompileResult> _variables = new Dictionary<string, CompileResult>();

        private string _variableName;

        internal void AddVariableName(string name)
        {
            _variableName = name;
        }

        internal bool VariableIsSet(string name)
        {
            return _variables.ContainsKey(name);
        }

        internal int NumberOfVariables => _variables.Count + (string.IsNullOrEmpty(_variableName) ? 0 : 1);

        internal void AddVariableValue(CompileResult value)
        {
            _variables[this._variableName] = value;
        }

        internal CompileResult GetVariableValue(string variableName)
        {
            if (_variables.ContainsKey(variableName))
            {
                return _variables[variableName];
            }
            return CompileResult.Empty;
        }
    }
}
