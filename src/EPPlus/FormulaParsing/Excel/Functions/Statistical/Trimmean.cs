/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "7.0",
        Description = "Returns the mean of the interior of a data set. TRIMMEAN calculates the mean taken by excluding a percentage of data points from the top and bottom tails of a data set. You can use this function when you wish to exclude outlying data from your analysis.")]
    internal class Trimmean : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var values = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments[0] }, context);
            var percentage = ArgToDecimal(arguments, 1, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            if (percentage < 0 || percentage >= 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }

            // cast ExcelDoubleValue to double
            var doubleValues = values.Select(x => (double)x);   
            var result = TrimMean(doubleValues.ToList(), percentage);
            return CreateResult(result, DataType.Decimal);
        }
        public static double TrimMean(List<double> values, double percentage)
        {

            values.Sort();

            int excludeCount = (int)Math.Round(values.Count * percentage);

            List<double> trimmedValues = values.Skip(excludeCount).Take(values.Count - 2 * excludeCount).ToList();

            double mean = trimmedValues.Average();

            return mean;
        }

    }
}
