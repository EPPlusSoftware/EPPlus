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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Sorting.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Statistical,
       EPPlusVersion = "7.2",
       Description = "The LINEST function calculates a regressional line that fits your data. It also calculates additional statistics." +
                     "It can handle several x-variables and perform multiple regression analysis.")]
    internal class Linest : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            //X can have more than one vector corresponding to each y-value
            if (!arguments[0].IsExcelRange) return CompileResult.GetErrorResult(eErrorType.Value);
            var constVar = true;
            var stats = false;
            if (arguments.Count() > 2 && arguments[2].DataType != DataType.Empty) constVar = ArgToBool(arguments, 2);
            if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);

            if (arguments.Count() > 1 && arguments[1].IsExcelRange)
            {
                var argY = arguments[0].ValueAsRangeInfo;
                var argX = arguments[1].ValueAsRangeInfo;
                var linestResult = LinestHelper.ExecuteLinest(argX, argY, constVar, stats, false, out eErrorType? error);
                if (error == null)
                {
                    return CreateDynamicArrayResult(linestResult, DataType.ExcelRange);
                }
                else
                {
                    return CreateResult(error.Value);
                }
            }
            else
            {
                var knownYsList = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments[0] }, context, out ExcelErrorValue e1).ToList();
                if (e1 != null) return CreateResult(e1.Type);
                var knownXs = LinestHelper.GetDefaultKnownXs(knownYsList.Count());
                var knownYs = MatrixHelper.ListToArray(knownYsList);
                var resultRange = LinestHelper.LinearRegResult(knownXs, knownYs, constVar, stats, false);

                return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
            }
        }
    }
}
