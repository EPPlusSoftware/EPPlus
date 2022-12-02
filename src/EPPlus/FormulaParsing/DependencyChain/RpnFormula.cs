using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing
{
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
        internal Stack<int> _funcStackPosition;
        internal CompileResult _currentResult;
        public RpnFormula(ExcelWorksheet ws, int row, int column)
        {
            _ws = ws;
            _row = row;
            _column = column;
            _expressionStack = new Stack<RpnExpression>();
            _funcStackPosition = new Stack<int>();
        }
        internal void SetFormula(string formula, ISourceCodeTokenizer tokenizer, RpnExpressionGraph graph)
        {
            _tokens = RpnExpressionGraph.CreateRPNTokens(tokenizer.Tokenize(formula));
            _formula= formula;
            _expressions = graph.CompileExpressions(ref _tokens);
        }
    }
}