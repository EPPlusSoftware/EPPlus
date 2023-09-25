/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "5.2",
        Description = "The Excel Percentrank function calculates the relative position, between 0 and 1 (inclusive), of a specified value within a supplied array.")]
    internal class Percentrank : RankFunctionBase
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var array = GetNumbersFromArgs(arguments, 0, context, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var number = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            if (number < array.First() || number > array.Last()) return CompileResult.GetErrorResult(eErrorType.NA);
            var significance = 3;
            if (arguments.Count > 2)
            {
                significance = ArgToInt(arguments, 2);
            }
            var result = PercentRankIncImpl(array, number);
            result = RoundResult(result, significance);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
