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

        public VariableStorageScope AddNewScope()
        {
            VariableStorageScope parent = _scopes.Count > 0 ? _scopes.Peek() : null;
            var newScope = new VariableStorageScope(this, parent);
            _scopes.Push(newScope);
            return newScope;
        }

        public void Clear()
        {
            _scopes.Clear();
        }

        public VariableStorageScope CurrentOrNew()
        {
            return _scopes.Count > 0 ? _scopes.Peek() : AddNewScope();
        }

        public VariableStorageScope Peek()
        {
            return _scopes.Peek();
        }

        public bool IsEmpty => _scopes == null || _scopes.Count == 0;
    }
}
