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
    internal class VariableStorageManager
    {
        public VariableStorageManager() { }

        private readonly Stack<VariableStorageScope> _scopes = new Stack<VariableStorageScope>();
        private readonly Dictionary<int, VariableStorageScope> _scopesById = new Dictionary<int, VariableStorageScope>();

        public VariableStorageScope AddNewScope()
        {
            VariableStorageScope parent = _scopes.Count > 0 ? _scopes.Peek() : null;
            var newScope = new VariableStorageScope(this, parent);
            _scopesById.Add(newScope.Id, newScope);
            _scopes.Push(newScope);
            return newScope;
        }

        public int Count => _scopes.Count;

        public void Clear()
        {
            _scopes.Clear();
            _scopesById.Clear();
        }

        public VariableStorageScope Peek()
        {
            return _scopes.Peek();
        }

        public VariableStorageScope GetById(int id)
        {
            if (!_scopesById.ContainsKey(id)) return null;
            return _scopesById[id];
        }

        public void Pop()
        {
            _scopes.Pop();
        }

        public void Push(VariableStorageScope scope)
        {
            _scopes.Push(scope);
            if(_scopesById.ContainsKey(scope.Id)) _scopesById.Remove(scope.Id); 
            _scopesById[scope.Id] = scope;  
        }

        public bool IsEmpty => _scopes == null || _scopes.Count == 0;
    }
}
