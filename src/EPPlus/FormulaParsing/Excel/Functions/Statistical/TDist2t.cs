using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    internal class TDist2t : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var x = ArgToDecimal(arguments, 0);
            var degreesOfFreedom = ArgToDecimal(arguments, 1);

            //Excel rounds degreesOfFreedom to the nearest int, so we do the same.

            degreesOfFreedom = System.Math.Floor(degreesOfFreedom);
            if (degreesOfFreedom < 1)
            {
                return CreateResult(eErrorType.Num);
            }

            if (x < 0)
            {
                return CreateResult(eErrorType.Num);
            }

            var result = 2 * (1 - StudenttHelper.CumulativeDistributionFuncion(x, degreesOfFreedom));

            return CreateResult(result, DataType.Decimal);
            //throw new NotImplementedException();
        }
    }
}
