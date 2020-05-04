using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    public class Bin2Dec : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToString(arguments, 0);
            if (number.Length > 10) return CreateResult(eErrorType.Num);
            if (number.Length < 10)
            {
                return CreateResult(Convert.ToInt32(number, 2), DataType.Decimal);
            }
            if (!BinaryHelper.TryParseBinaryToDecimal(number, 2, out int result)) return CreateResult(eErrorType.Num);
            return CreateResult(result, DataType.Integer);
        }
    }
}
