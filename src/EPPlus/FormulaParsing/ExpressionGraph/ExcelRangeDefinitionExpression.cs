using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class ExcelRangeDefinitionExpression : Expression
    {
        FormulaRangeAddress _range;
        internal ExcelRangeDefinitionExpression(FormulaRangeAddress range, ParsingContext ctx)
            : base(ctx)
        {
            _range = range;
        }

        public override bool IsGroupedExpression => false;

        internal override ExpressionType ExpressionType => ExpressionType.RangeDefinition;

        public override CompileResult Compile()
        {
            var dp = new EpplusExcelDataProvider(Context.Package, Context);
            return new CompileResult(dp.GetRange(_range), DataType.Enumerable);
        }

        internal override Expression Clone()
        {
            return CloneMe(new ExcelRangeDefinitionExpression(_range, Context));
        }
    }
}
