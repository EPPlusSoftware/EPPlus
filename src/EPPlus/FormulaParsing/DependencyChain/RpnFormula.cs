using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing
{
    internal class RpnFunction        
    {
        public RpnFunction(int startPos, string function)
        {
            _startPos = startPos;
            _function = function;
        }
        internal int _startPos;
        internal string _function;
        internal List<short> _arguments;
    }
    internal class RpnFormula
    {
        internal ExcelWorksheet _ws;
        internal int _row;
        internal int _column;
        internal string _formula;
        internal IList<Token> _tokens;
        internal Dictionary<int, RpnExpression> _expressions;
        internal CellStoreEnumerator<object> _formulaEnumerator;
        internal int _tokenIndex = 0;
        internal Stack<RpnExpression> _expressionStack;
        internal Stack<RpnFunctionExpression> _funcStack;
        internal int _arrayIndex = -1;
        internal bool _isDynamic = false;
        internal RpnFormula(ExcelWorksheet ws, int row, int column)
        {
            _ws = ws;
            _row = row;
            _column = column;
            _expressionStack = new Stack<RpnExpression>();
            _funcStack = new Stack<RpnFunctionExpression>();
        }

        internal string GetAddress()
        {
            
            if(_ws==null)
            {
                return ExcelCellBase.GetAddress(_row, _column);
            }
            return _ws.Name + "!" + ExcelCellBase.GetAddress(_row, _column);
        }

        internal void SetFormula(string formula, RpnOptimizedDependencyChain depChain)
        {
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(_ws==null ? -1 : _ws.IndexInList, _row, _column);
            _tokens = RpnExpressionGraph.CreateRPNTokens(depChain._tokenizer.Tokenize(formula));
            _formula= formula;
            _expressions = RpnExpressionGraph.CompileExpressions(ref _tokens, depChain._parsingContext);
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
    }
}