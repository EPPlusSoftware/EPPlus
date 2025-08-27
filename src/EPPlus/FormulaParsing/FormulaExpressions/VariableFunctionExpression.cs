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
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class VariableFunctionExpression : FunctionExpression
    {


        internal VariableFunctionExpression(string tokenValue, ParsingContext ctx, int pos, bool addVariableScope = true) : base(tokenValue, ctx, pos)
        {
            _variableStorage = ctx.VariableStorage;
            if (addVariableScope)
            {
                _storageScope = _variableStorage.AddNewScope();
                _isOutOfVariableScope = false;
            }
            else if(!_variableStorage.IsEmpty)
            {
                _storageScope = _variableStorage.Peek();
                _isOutOfVariableScope = false;
            }
            else
            {
                _isOutOfVariableScope = true;
            }
             VariableScopeId = _isOutOfVariableScope ? -1 : _storageScope.Id;
        }

        private string _lastDeclaredVariable;
        private readonly VariableStorageManager _variableStorage;
        private readonly VariableStorageScope _storageScope;
        private bool _isOutOfVariableScope = false;


        internal int VariableScopeId { get; private set; }

        internal VariableStorageScope VariableScope => _storageScope;

        internal override void OnExecuteStarted()
        {
            
        }

        internal override bool IsVariable(string name)
        {
            return VariableIsSet(name);
        }

        internal void DeclareVariable(string name)
        {
            if (_isOutOfVariableScope) return;
            if (!_storageScope.ContainsVariable(name))
            {
                _storageScope.SetVariableValue(name, null);
            }
            _lastDeclaredVariable = name;
        }

        internal bool VariableIsDeclared(string name)
        {
            if (_isOutOfVariableScope) return false;
            if (_storageScope.ContainsVariable(name))
            {
                return true;
            }
            return false;
        }

        internal bool VariableIsSet(string name)
        {
            if (_isOutOfVariableScope) return false;
            if(_storageScope.ContainsVariable(name) && _storageScope.GetVariableValue(name) != null)
            {
                return true;
            }
            return false;
        }

        internal int NumberOfVariables => _storageScope.NumberOfVariables;

        internal void AddVariableValue(CompileResult value)
        {
            if (_isOutOfVariableScope) return;
            _storageScope.SetVariableValue(_lastDeclaredVariable, value);
        }

        internal void AddVariableValue(string name, CompileResult value)
        {
            if (_isOutOfVariableScope) return;
            _storageScope.SetVariableValue(name, value);
        }

        internal CompileResult GetVariableValue(string variableName)
        {
            if(_isOutOfVariableScope)
            {
                return new CompileResult("oovs", DataType.Empty);
            }
            if (_storageScope.ContainsVariable(variableName))
            {
                return _storageScope.GetVariableValue(variableName) ?? CompileResult.Empty;
            }
            return CompileResult.Empty;
        }
    }
}
