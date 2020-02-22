using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class Unichar : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            if (
                IsNumeric(arguments.ElementAt(0).Value)
                &&
                short.TryParse(ArgToString(arguments, 0), out short arg))
            {
                return CreateResult(char.ConvertFromUtf32(arg), DataType.Integer);
            }
            return CreateResult(ExcelErrorValue.Values.Value, DataType.ExcelError);
        }
    }
}
