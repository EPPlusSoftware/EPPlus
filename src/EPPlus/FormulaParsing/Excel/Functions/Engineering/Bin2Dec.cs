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
            //var result = Convert.ToInt32(number, 2);
            if (number.Length < 10)
            {
                return CreateResult(Convert.ToInt32(number, 2), DataType.Decimal);
            }
            var chars = number.ToCharArray();
            var isNegative = chars[0] == '1';
            var negativeUsed = false;
            var result = 0;
            for(var x = 1; x < 10; x++)
            {
                var c = chars[x];
                var current = 0;
                if (c != '0' && c != '1') return CreateResult(eErrorType.Num);
                if(x == 9)
                {
                    current = c == '1' ? 1 : 0;
                    if (isNegative && !negativeUsed) current *= -1;
                }
                else if (c == '1')
                {
                    current = (int)System.Math.Pow(2, 9 - x);
                    if (isNegative && !negativeUsed) current *= -1;
                    negativeUsed = true;
                }
                result += current;
            }
            return CreateResult(result, DataType.Integer);
        }
    }
}
