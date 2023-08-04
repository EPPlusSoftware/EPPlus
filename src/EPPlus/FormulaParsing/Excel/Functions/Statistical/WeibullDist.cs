/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/07/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{

    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the Weibull distribution. This function works the same as WEIBULL.")]
    internal class WeibullDist : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 4;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var x = ArgToDecimal(arguments, 0);
            var alpha = ArgToDecimal(arguments, 1);
            var beta = ArgToDecimal(arguments, 2);
            var cumulative = ArgToBool(arguments, 3);
            var weibull = 0d;
            
            if (x < 0 || alpha <= 0 || beta <= 0)
            {
                return CreateResult(eErrorType.Num);
            }

            if (alpha == 1)
            {
                return CreateResult(expDistribution(x, 1 / beta, cumulative), DataType.Decimal);
            }

            if (cumulative)
            {
                weibull = 1 - Math.Exp(-Math.Pow((x / beta), alpha));
            }
            else
            {
                weibull = alpha / Math.Pow(beta, alpha) * Math.Pow(x, alpha - 1) * Math.Exp(-Math.Pow(x / beta, alpha));
            }

            return CreateResult(weibull, DataType.Decimal);
        }

        internal static double expDistribution(double x, double lambda, bool cumulative)
        {
            var result = 0d;
            if (cumulative)
            {
                result = 1d - Math.Exp(x * -lambda);
            }
            else
            {
                result = lambda * Math.Exp(x * -lambda);
            }
            return result;
        }
    }
}
