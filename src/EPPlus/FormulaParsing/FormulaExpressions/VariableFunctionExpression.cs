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
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
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


        internal VariableFunctionExpression(string tokenValue, VariableStorageManager variableStorage, ParsingContext ctx, int pos) : base(tokenValue, ctx, pos)
        {
            _variableStorage = variableStorage;
            _storageScope = _variableStorage.AddNewScope();
            VariableScopeId = _storageScope.Id;
        }

        private readonly Dictionary<string, CompileResult> _variables = new Dictionary<string, CompileResult>();
        private string _lastDeclaredVariable;
        private readonly VariableStorageManager _variableStorage;
        private readonly VariableStorageScope _storageScope;


        internal int VariableScopeId { get; private set; }

        internal override void OnExecuteStarted()
        {
            
        }
        internal override bool IsVariable(string name)
        {
            return VariableIsSet(name);
        }

        internal void DeclareVariable(string name)
        {
            if (!_storageScope.ContainsVariable(name))
            {
                _storageScope.SetVariableValue(name, null);
            }
            _lastDeclaredVariable = name;
        }

        internal bool VariableIsDeclared(string name)
        {
            if (_storageScope.ContainsVariable(name))
            {
                return true;
            }
            return false;
        }

        internal bool VariableIsSet(string name)
        {
            if(_storageScope.ContainsVariable(name) && _storageScope.GetVariableValue(name) != null)
            {
                return true;
            }
            return false;
        }

        internal int NumberOfVariables => _storageScope.NumberOfVariables;

        internal void AddVariableValue(CompileResult value)
        {
            _storageScope.SetVariableValue(_lastDeclaredVariable, value);
        }

        internal void AddVariableValue(string name, CompileResult value)
        {
            _storageScope.SetVariableValue(name, value);
        }

        internal CompileResult GetVariableValue(string variableName)
        {
            if (_storageScope.ContainsVariable(variableName) && _storageScope.GetVariableValue(variableName) != null)
            {
                return _storageScope.GetVariableValue(variableName);
            }
            return CompileResult.Empty;
        }
    }
}
