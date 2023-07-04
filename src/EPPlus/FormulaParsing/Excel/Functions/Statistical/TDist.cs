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

            //Based on out tests, degrees of freedom is rounded down to the nearest integer when input is a decimal.
            degreesOfFreedom = System.Math.Floor(degreesOfFreedom);

            if (degreesOfFreedom < 1)
            {
                return CreateResult(eErrorType.Div0);
            }

            if (cumulative)
            {
                var result = StudenttHelper.CumulativeDistributionFuncion(x, degreesOfFreedom);
                return CreateResult(result, DataType.Decimal);
            }
            else
            {
                var result = StudenttHelper.ProbabilityDensityFunction(x, degreesOfFreedom);
                return CreateResult(result, DataType.Decimal);
            }

        }

    }
}
