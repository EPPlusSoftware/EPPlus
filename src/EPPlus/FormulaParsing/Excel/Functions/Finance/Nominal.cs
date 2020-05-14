using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    internal class Nominal : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var effectRate = ArgToDecimal(arguments, 0);
            var npery = ArgToInt(arguments, 1);
            if (effectRate <= 0 || npery < 1)
                return CreateResult(eErrorType.Num);
            var result = (System.Math.Pow(effectRate + 1d, 1d / npery) - 1d) * npery;
            return CreateResult(result, DataType.Decimal);
        }
    }
}
