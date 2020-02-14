using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class Trim : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var str = ArgToString(arguments, 0);
            if(!string.IsNullOrEmpty(str))
            {
                str = str.Trim();
                str = Regex.Replace(str, "[ ]+", " ");
            }
            return CreateResult(str, DataType.String);
        }
    }
}
