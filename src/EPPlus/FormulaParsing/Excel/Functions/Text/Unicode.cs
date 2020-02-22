using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class Unicode : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = ArgToString(arguments, 0);
            if (!IsString(arg, allowNullOrEmpty: false)) return CreateResult(ExcelErrorValue.Values.Value, DataType.ExcelError);
            var firstChar = arg.Substring(0, 1);
            var bytes = Encoding.UTF32.GetBytes(firstChar);
            var code = BitConverter.ToInt32(bytes, 0);
            return CreateResult(code, DataType.Integer);
        }
    }
}
