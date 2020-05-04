using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    public class Bin2Hex : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToString(arguments, 0);
            var formatString = "X";
            if(arguments.Count() > 1)
            {
                var padding = ArgToInt(arguments, 1);
                if (padding < 0 ^ padding > 10) return CreateResult(eErrorType.Num);
                formatString += padding;
            }
            if (number.Length > 10) return CreateResult(eErrorType.Num);
            if (number.Length < 10)
            {
                var n = Convert.ToInt32(number, 2);
                return CreateResult(n.ToString(formatString), DataType.Decimal);
            }
            else
            {
                if (!BinaryHelper.TryParseBinaryToDecimal(number, 2, out int result)) return CreateResult(eErrorType.Num);
                var hexStr = result.ToString(formatString);
                hexStr = BinaryHelper.EnsureLength(hexStr, 10, "F");
                return CreateResult(hexStr, DataType.String);
            }
        }
    }
}
