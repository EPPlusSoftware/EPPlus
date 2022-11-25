using OfficeOpenXml.Core.CellStore;
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
        internal IList<RpnExpression> _expressions;
        internal CellStoreEnumerator<object> _formulaEnumerator;
        internal int _expressionIndex = 0;
        internal Stack<RpnExpression> _expressionStack;
        internal Stack<int> _funcStackPosition;

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
            var tokens =  graph.CreateExpressionList(tokenizer.Tokenize(formula));
            _formula= formula;
            _expressions = graph.CompileExpressions(tokens);
        }
    }
}