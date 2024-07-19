/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 19/06/2024         EPPlus Software AB       Initial release EPPlus 7
*************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Sorting.Internal;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Statistical,
       EPPlusVersion = "7.0",
       Description = "The LOGEST function calculates an exponential curve that best fits the input data. It can also provide multiple curves if there are multiple " +
                     "x-variables.")]
    internal class Logest : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!arguments[0].IsExcelRange) return CompileResult.GetErrorResult(eErrorType.Value);
            var constVar = true;
            var stats = false;
            if (arguments.Count() > 2 && arguments[2].DataType != DataType.Empty) constVar = ArgToBool(arguments, 2);
            if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);

            if (arguments.Count() > 1 && arguments[1].IsExcelRange)
            {
                var rangeX = arguments[1].ValueAsRangeInfo;
                var rangeY = arguments[0].ValueAsRangeInfo;
                var linestResult = LinestHelper.ExecuteLinest(rangeX, rangeY, constVar, stats, true, out eErrorType? error);
                if (error == null)
                {
                    return CreateDynamicArrayResult(linestResult, DataType.ExcelRange);
                }
                else
                {
                    return CompileResult.GetErrorResult(error.Value);
                }
            }
            else
            {
                var knownYsList = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments[0] }, context, out ExcelErrorValue e1).ToList();
                if (e1 != null) return CreateResult(e1.Type);
                var knownXs = LinestHelper.GetDefaultKnownXs(knownYsList.Count());
                var knownYs = MatrixHelper.ListToArray(knownYsList);
                var resultRange = LinestHelper.LinearRegResult(knownXs, knownYs, constVar, stats, true);

                return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
            }
        }
    }
}
