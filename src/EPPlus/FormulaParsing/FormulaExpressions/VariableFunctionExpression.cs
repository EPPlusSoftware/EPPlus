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


        internal void AddVariableName(string name)
        {
            if (!_variables.ContainsKey(name))
            {
                _variables.Add(name, null);
            }
        }

        internal bool VariableIsSet(string name)
        {
            return _variables.ContainsKey(name);
        }

        internal int NumberOfVariables => _variables.Count;

        internal void AddVariableValue(string name, CompileResult value)
        {
            _variables[name] = value;
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
