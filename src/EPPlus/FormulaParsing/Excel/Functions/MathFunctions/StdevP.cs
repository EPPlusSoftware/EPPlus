/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MathObj = System.Math;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the standard deviation of a supplied set of values (which represent an entire population)")]
    internal class StdevP : HiddenValuesHandlingFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public StdevP()
        {
            IgnoreErrors = false;
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var args = ArgsToDoubleEnumerable(arguments, context, x =>
            {
                x.IgnoreHiddenCells = IgnoreHiddenValues;
                x.IgnoreErrors = IgnoreErrors;
                x.IgnoreNestedSubtotalAggregate = IgnoreNestedSubtotalsAndAggregates;
            }, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            return CreateResult(StandardDeviation(args), DataType.Decimal);
        }

        internal static double StandardDeviation(IEnumerable<double> values)
        {
            double avg = values.AverageKahan();
            return MathObj.Sqrt(values.AverageKahan(v => MathObj.Pow(v - avg, 2)));
        }
    }
}
