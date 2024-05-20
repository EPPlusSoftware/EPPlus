/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/14/2024         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/
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

        internal override bool IsVariable(string name)
        {
            return VariableIsSet(name);
        }

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
