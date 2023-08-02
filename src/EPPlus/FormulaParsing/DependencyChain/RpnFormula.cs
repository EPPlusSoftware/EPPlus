﻿using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Security.AccessControl;

namespace OfficeOpenXml.FormulaParsing
{
    internal class RpnFormula
    {
        internal ExcelWorksheet _ws;
        internal int _row;
        internal int _column;
        internal string _formula;
        internal IList<Token> _tokens;
        internal Dictionary<int, Expression> _expressions;
        internal CellStoreEnumerator<object> _formulaEnumerator;
        internal int _tokenIndex = 0;
        internal Stack<Expression> _expressionStack;
        internal Stack<FunctionExpression> _funcStack;
        internal int _arrayIndex = -1;
        internal bool _isDynamic = false;

        public bool CanBeDynamicArray 
        {
            get
            {
                return _ws._flags.GetFlagValue(_row, _column, CellFlags.CanBeDynamicArray);
            }
        }

        internal RpnFormula(ExcelWorksheet ws, int row, int column)
        {
            _ws = ws;
            _row = row;
            _column = column;
            _expressionStack = new Stack<Expression>();
            _funcStack = new Stack<FunctionExpression>();
        }

        internal string GetAddress()
        {
            
            if(_ws==null)
            {
                if(_row>=0 && _column>=0)
                {
                    return ExcelCellBase.GetAddress(_row, _column);
                }
                else
                {
                    return $"Workbook name - index {_row}";
                }
            }
            return _ws.Name + "!" + ExcelCellBase.GetAddress(_row, _column);
        }

        internal void SetFormula(string formula, RpnOptimizedDependencyChain depChain)
        {
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(_ws==null ? -1 : _ws.IndexInList, _row, _column);
            _tokens = FormulaExecutor.CreateRPNTokens(depChain._tokenizer.Tokenize(formula));
            _formula= formula;
            _expressions = FormulaExecutor.CompileExpressions(ref _tokens, depChain._parsingContext);
        }
        public override string ToString()
        {
            if(_ws==null)
            {
                return ExcelCellBase.GetAddress(_row, _column);
            }
            else
            {
                return _ws.Name + "!" + ExcelCellBase.GetAddress(_row, _column);
            }
        }

        internal void ClearCache()
        {
            foreach(var e in _expressions.Values)
            {
                e._cachedCompileResult = null;
            }
        }
    }
    internal class RpnNameFormula : RpnFormula
    {        
        internal RpnNameFormula(ExcelWorksheet ws, int nameId, int row, int column) : base(ws, row, column)
        {
            NameId  = nameId;

        }
        internal int NameId { get; set; }
        internal bool IsStatic { get; set; }
        internal object Value { get; set; }
    }
}