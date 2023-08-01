/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/07/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the left-tailed Students t-distribution. The Students t-distribution is used for hypothesis testing with small samples.")]
    internal class TDist : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
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
                var result = StudenttHelper.CumulativeDistributionFunction(x, degreesOfFreedom);
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
