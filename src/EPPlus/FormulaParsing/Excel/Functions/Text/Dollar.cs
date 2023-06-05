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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Text,
       EPPlusVersion = "5.5",
       Description = "Converts a supplied number into text, using a currency format")]
    internal class Dollar : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var number = ArgToDecimal(arguments, 0, context.Configuration.PrecisionAndRoundingStrategy);
            var decimals = 2;
            if(arguments.Count > 1)
            {
                decimals = ArgToInt(arguments, 1);
            }
            double result;
            if(decimals >= 0)
            {
                result = Math.Round(number, decimals);
            }
            else
            {
                result = Math.Round(number * System.Math.Pow(10, decimals)) / System.Math.Pow(10, decimals);
            }
            return CreateResult(result.ToString(GetFormatString(decimals), CultureInfo.CurrentCulture), DataType.String);
        }

        private string GetFormatString(int decimals)
        {
            if (decimals > 0) return "C" + decimals;
            return "C0";
        }
    }
}
