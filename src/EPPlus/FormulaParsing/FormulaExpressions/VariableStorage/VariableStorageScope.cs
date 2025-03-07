/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/12/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage
{
    internal class VariableStorageScope
    {
        public VariableStorageScope(VariableStorageManager storageManager)
            : this(storageManager, null)
        {
            
        }
        public VariableStorageScope(VariableStorageManager storageManager, VariableStorageScope parentScope)
        {
            _parentScope = parentScope;
            Id = VariableStorageId.GetNewId();
            _storageManager = storageManager;
        }

        private readonly VariableStorageScope _parentScope;
        private readonly VariableStorageManager _storageManager;

        private readonly Dictionary<string, CompileResult> _variables = new Dictionary<string, CompileResult>();

        public int Id { get; private set; }

        public VariableStorageManager VariableStorage => _storageManager;

        public int NumberOfVariables => _variables.Count + _parentScope?.NumberOfVariables ?? 0;



        public bool ContainsVariable(string name)
        {
            if(_variables.ContainsKey(name)) return true;
            if (_parentScope == null) return false;
            return _parentScope.ContainsVariable(name);
        }

        public CompileResult GetVariableValue(string name)
        {
            if( _variables.ContainsKey(name))
                return _variables[name];
            if (_parentScope.ContainsVariable(name))
                return _parentScope.GetVariableValue(name);
            return CompileResult.Empty;
        }

        public void SetVariableValue(string name, CompileResult value)
        {
            if(_variables.ContainsKey(name)) _variables.Remove(name);
            _variables[name] = value;
        }
    }
}
