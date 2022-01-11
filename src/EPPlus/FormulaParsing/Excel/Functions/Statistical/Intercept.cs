using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "6.0",
        Description = "Calculates the point at which a line will intersect the y-axis by using existing x-values and y-values.")]
    internal class Intercept : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arg1 = arguments.ElementAt(0);
            var arg2 = arguments.ElementAt(1);
            var knownYs = ArgsToDoubleEnumerable(false, false, new FunctionArgument[] { arg1 }, context).Select(x => x.Value).ToArray();
            var knownXs = ArgsToDoubleEnumerable(false, false, new FunctionArgument[] { arg2 }, context).Select(x => x.Value).ToArray();
            if (knownYs.Count() != knownXs.Count()) return CreateResult(eErrorType.NA);
            if (!knownYs.Any()) return CreateResult(eErrorType.NA);
            var result = InterceptImpl(0, knownYs, knownXs);
            return CreateResult(result, DataType.Decimal);
        }

        internal static double InterceptImpl(double x, double[] arrayY, double[] arrayX)
        {
            var avgY = arrayY.Average();
            var avgX = arrayX.Average();
            var nItems = arrayY.Length;
            var upperEquationPart = 0d;
            var lowerEquationPart = 0d;
            for (var ix = 0; ix < nItems; ix++)
            {
                upperEquationPart += (arrayX[ix] - avgX) * (arrayY[ix] - avgY);
                lowerEquationPart += System.Math.Pow(arrayX[ix] - avgX, 2);
            }
            return avgY - (upperEquationPart / lowerEquationPart) * avgX;
        }
    }
}
