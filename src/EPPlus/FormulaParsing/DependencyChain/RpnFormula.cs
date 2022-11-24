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
        internal CellStoreEnumerator<CellStoreValue> _formulaEnumerator;
        internal int _addressExpressionIndex;
        public RpnFormula(ExcelWorksheet ws, int row, int column)
        {
            _ws = ws;
            _row = row;
            _column = column;
        }
        internal void SetFormula(string formula, ISourceCodeTokenizer tokenizer, RpnExpressionGraph graph)
        {
            var tokens =  graph.CreateExpressionList(tokenizer.Tokenize(formula));
            var _expresions = graph.CompileExpressions(tokens);
        }
    }
}