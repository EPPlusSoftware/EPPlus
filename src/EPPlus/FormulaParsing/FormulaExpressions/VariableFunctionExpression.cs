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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class VariableFunctionExpression : FunctionExpression
    {


        internal VariableFunctionExpression(string tokenValue, Stack<FunctionExpression> funcStack, ParsingContext ctx, int pos) : base(tokenValue, ctx, pos)
        {

        }

        private readonly Dictionary<string, CompileResult> _variables = new Dictionary<string, CompileResult>();
        private string _lastDeclaredVariable;

        internal override bool IsVariable(string name)
        {
            return VariableIsSet(name);
        }

        internal void DeclareVariable(string name)
        {
            if (!_variables.ContainsKey(name))
            {
                _variables.Add(name, null);
            }
            _lastDeclaredVariable = name;
        }

        internal bool VariableIsDeclared(string name)
        {
            if (_variables.ContainsKey(name))
            {
                return true;
            }
            return false;
        }

        internal bool VariableIsSet(string name)
        {
            if(_variables.ContainsKey(name) && _variables[name] !=null)
            {
                return true;
            }
            return false;
        }

        internal int NumberOfVariables => _variables.Count;

        internal void AddVariableValue(CompileResult value)
        {
            _variables[_lastDeclaredVariable] = value;
        }

        internal void AddVariableValue(string name, CompileResult value)
        {
            _variables[name] = value;
        }

        internal CompileResult GetVariableValue(string variableName)
        {
            if (_variables.ContainsKey(variableName) && _variables[variableName] != null)
            {
                return _variables[variableName];
            }
            return CompileResult.Empty;
        }
    }
}
