using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
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
            Description = "Calculates the inverse of the right-tailed probability of the Chi-Square Distribution. Same implementation as CHISQ.INV.RT")]
    internal class ChiInv : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var n = ArgToDecimal(arguments, 0);
            var degreesOfFreedom = ArgToInt(arguments, 1);
            if (n < 0d || degreesOfFreedom < 1 || degreesOfFreedom > System.Math.Pow(10, 10))
            {
                return CreateResult(eErrorType.Num);
            }
            var result = ChiSquareHelper.Inverse(1d - n, degreesOfFreedom);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
