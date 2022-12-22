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

        internal void SetFormula(string formula, ISourceCodeTokenizer tokenizer, RpnExpressionGraph graph)
        {
            _tokens = RpnExpressionGraph.CreateRPNTokens(tokenizer.Tokenize(formula));
            _formula= formula;
            _expressions = graph.CompileExpressions(ref _tokens);
        }
    }
}