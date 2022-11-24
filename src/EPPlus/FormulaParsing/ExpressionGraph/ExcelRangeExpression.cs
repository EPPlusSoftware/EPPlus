using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExcelRangeExpression : Expression
    {
        public ExcelRangeExpression(IRangeInfo rangeInfo, ParsingContext ctx)
            : base(ctx)
        {
            _rangeInfo = rangeInfo;
        }

        private readonly IRangeInfo _rangeInfo;
        public override bool IsGroupedExpression => false;
        internal override ExpressionType ExpressionType => ExpressionType.ExcelRange;
        public override CompileResult Compile()
        {
            return new CompileResult(_rangeInfo, DataType.ExcelRange);
        }

        internal override Expression Clone()
        {
            return CloneMe(new ExcelRangeExpression(_rangeInfo, Context));
        }
    }
}
