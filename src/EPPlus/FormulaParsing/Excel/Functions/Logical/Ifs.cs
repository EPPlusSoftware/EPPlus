using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    public class Ifs : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var crf = new CompileResultFactory();
            var maxArgs = arguments.Count() < (127 * 2) ? arguments.Count() : 127 * 2; 
            for(var x = 0; x < maxArgs; x += 2)
            {
                if (System.Math.Round(ArgToDecimal(arguments, x), 15) != 0d) return crf.Create(arguments.ElementAt(x + 1).Value);
            }
            return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
        }
    }
}
