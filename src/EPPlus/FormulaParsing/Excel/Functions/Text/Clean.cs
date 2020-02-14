using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class Clean : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var str = ArgToString(arguments, 0);
            if(!string.IsNullOrEmpty(str))
            {
                var sb = new StringBuilder();
                var arr = Encoding.ASCII.GetBytes(str);
                foreach(var c in arr)
                {
                    if (c > 31)
                        sb.Append((char)c);
                }
                str = sb.ToString();
            }
            return CreateResult(str, DataType.String);
        }
    }
}
