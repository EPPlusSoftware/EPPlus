using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal class RangeOffset : ExcelFunction
    {
        public string StartRange { get; set; }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            // EXAMPLE: A1:OFFSET(B2, 2, 0)

            var func = new Offset();
            // this call is OFFSET(B2, 2, 0)
            var offsetRangeResult = func.Execute(arguments, context);
            // Here is the result from the OFFSET function
            var offsetRange = offsetRangeResult.Result as ExcelDataProvider.IRangeInfo;
            if (offsetRange == null) return CreateResult(eErrorType.Value);

            // A1 should be set as StartRange by the UnregognizedFunctionName pipline in the FunctionExpression.

            // TODO: create a new range from these two ranges according to how Excel does it
            //       return this range like this:
            // return CreateResult(range, DataType.Enumerable);

            throw new NotImplementedException();
        }
    }
}
