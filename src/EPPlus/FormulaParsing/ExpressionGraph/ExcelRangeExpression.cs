using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExcelRangeExpression : Expression
    {
        public ExcelRangeExpression(ExcelDataProvider.IRangeInfo rangeInfo)
        {
            _rangeInfo = rangeInfo;
        }

        private readonly ExcelDataProvider.IRangeInfo _rangeInfo;
        public override bool IsGroupedExpression => false;

        public override CompileResult Compile(bool treatEmptyAsZero = true)
        {
            return new CompileResult(_rangeInfo, DataType.Enumerable);
        }
    }
}
