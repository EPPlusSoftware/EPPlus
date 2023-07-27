/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/07/2023         EPPlus Software AB         Implemented function
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
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
    Description = "Calculates the result of the F-test")]
    internal class FTest : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var array1 = arguments[0].ValueAsRangeInfo;
            var array2 = arguments[1].ValueAsRangeInfo;

            RangeFlattener.GetNumericPairLists(array1, array2, false, out List<double> list1, out List<double> list2);
            if (list1.Count() < 2 || list2.Count() < 2) return CreateResult(eErrorType.Div0);

            double variance1 = VarianceCalc(list1);
            double variance2 = VarianceCalc(list2);

            if (variance1 == 0 || variance2 == 0) return CreateResult(eErrorType.Div0);

            var fStatistics = (variance2 > variance1) ? variance1 / variance2 : variance2 / variance1; //The smallest variance is divided by the largest.
            var df1 = list1.Count() - 1;
            var df2 = list2.Count() - 1;
            var result = 2 * FHelper.GetProbability(fStatistics, df1, df2, true);

            return CreateResult(result, DataType.Decimal);
        }

        internal static double VarianceCalc(List<double> values)
        {
            var mean = values.Average();
            var sumOfSquares = values.Sum(val => Math.Pow(val - mean, 2));
            var setVariance = sumOfSquares / (values.Count() - 1);
            return setVariance;
        }
    }
}

