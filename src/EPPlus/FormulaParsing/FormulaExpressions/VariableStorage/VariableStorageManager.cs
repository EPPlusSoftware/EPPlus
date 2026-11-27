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

        // Compilation boundaries (issue #2606). CompileExpressions can be re-entered at runtime
        // (e.g. a lambda body compiled while an outer formula is being calculated), so these are stacks.
        // Scopes at or below the recorded depth belong to the runtime/caller and must never become
        // lexical parents of scopes created during the compilation.
        private readonly Stack<int> _compileBaseCounts = new Stack<int>();
        private readonly Stack<VariableStorageScope> _compileRoots = new Stack<VariableStorageScope>();

        /// <summary>
        /// Starts a compilation. <paramref name="root"/> is the lexical parent for scopes created
        /// at the top level of the compiled tokens (the lambda scope when compiling a lambda body, otherwise null).
        /// </summary>
        /// <param name="root">The lexical root scope, or null</param>
        public void BeginCompile(VariableStorageScope root)
        {
            _compileBaseCounts.Push(_scopes.Count);
            _compileRoots.Push(root);
        }

        /// <summary>
        /// Pops scopes until the stack has the given depth. Does nothing if the stack is already at or below the depth.
        /// Popped scopes remain registered by id.
        /// </summary>
        /// <param name="depth">The depth to restore to</param>
        public void RestoreDepth(int depth)
        {
            while (_scopes.Count > depth)
            {
                _scopes.Pop();
            }
        }

        /// <summary>
        /// Ends the current compilation started with <see cref="BeginCompile(VariableStorageScope)"/>
        /// and restores the stack to the depth it had when the compilation started.
        /// </summary>
        public void EndCompile()
        {
            if (_compileBaseCounts.Count > 0)
            {
                var baseCount = _compileBaseCounts.Pop();
                _compileRoots.Pop();
                RestoreDepth(baseCount);
            }
        }

        /// <summary>
        /// Ends a top-level calculation. Restores the stack to the depth it had when the calculation started.
        /// When the outermost calculation ends, the id registry is cleared to avoid unbounded growth
        /// in long-lived packages.
        /// </summary>
        /// <param name="depth">The stack depth when the calculation started</param>
        public void EndCalculation(int depth)
        {
            RestoreDepth(depth);
            if (depth == 0 && _compileBaseCounts.Count == 0)
            {
                _scopesById.Clear();
            }
        }

        /// <summary>
        /// Returns the lexically enclosing scope. During compilation this is the innermost scope
        /// pushed by the current compilation, or the compilation root if none has been pushed.
        /// Outside compilation it is the top of the stack, or null if the stack is empty.
        /// </summary>
        public VariableStorageScope PeekLexical()
        {
            if (_compileBaseCounts.Count > 0 && _scopes.Count <= _compileBaseCounts.Peek())
            {
                return _compileRoots.Peek();
            }
            return _scopes.Count > 0 ? _scopes.Peek() : null;
        }

        public VariableStorageScope AddNewScope()
        {
            VariableStorageScope parent = PeekLexical();
            var newScope = new VariableStorageScope(this, parent);
            _scopesById.Add(newScope.Id, newScope);
            _scopes.Push(newScope);
            return newScope;
        }

        public int Count => _scopes.Count;

        public void Clear()
        {
            // Compilation boundaries are intentionally not cleared here,
            // they are owned by BeginCompile/EndCompile.
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
            if (_scopesById.ContainsKey(scope.Id)) _scopesById.Remove(scope.Id);
            _scopesById[scope.Id] = scope;
        }

        public bool IsEmpty => _scopes == null || _scopes.Count == 0;
    }
}
