using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    public class Bin2Oct : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToString(arguments, 0);
            var padding = default(int?);
            if(arguments.Count() > 1)
            {
                padding = ArgToInt(arguments, 1);
                if (padding.Value < 0 ^ padding.Value > 10) return CreateResult(eErrorType.Num);
            }
            if (number.Length > 10) return CreateResult(eErrorType.Num);
            var octStr = string.Empty;
            if (number.Length < 10)
            {
                var n = Convert.ToInt32(number, 2);
                octStr = Convert.ToString(n, 8);
            }
            else
            {
                if (!BinaryHelper.TryParseBinaryToDecimal(number, 2, out int result)) return CreateResult(eErrorType.Num);
                octStr = Convert.ToString(result, 8);
            }
            if(padding.HasValue)
            {
                octStr = BinaryHelper.EnsureLength(octStr, 10, "0");
            }
            return CreateResult(octStr, DataType.String);
        }
    }
}
