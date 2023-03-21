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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "5.2",
        Description = "The Excel Percentrank function calculates the relative position, between 0 and 1 (inclusive), of a specified value within a supplied array.")]
    internal class Percentrank : RankFunctionBase
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var array = GetNumbersFromArgs(arguments, 0, context);
            var number = ArgToDecimal(arguments, 1);
            if (number < array.First() || number > array.Last()) return CreateResult(eErrorType.NA);
            var significance = 3;
            if (arguments.Count() > 2)
            {
                significance = ArgToInt(arguments, 2);
            }
            var result = PercentRankIncImpl(array, number);
            result = RoundResult(result, significance);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
