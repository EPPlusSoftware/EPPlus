using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    internal class TDist : ExcelFunction
    {
        public override int ArgumentMinLength => 3;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var x = ArgToDecimal(arguments, 0);
            var degreesOfFreedom = ArgToDecimal(arguments, 1);
            var cumulative = ArgToBool(arguments, 2);

            if (cumulative)
            {
                var result = StudenttHelper.CDF(x, degreesOfFreedom);
                return CreateResult(result, DataType.Decimal);
            }
            else
            {
                var result = StudenttHelper.PDF(x, degreesOfFreedom);
                return CreateResult(result, DataType.Decimal);
            }

        }
    }
}
