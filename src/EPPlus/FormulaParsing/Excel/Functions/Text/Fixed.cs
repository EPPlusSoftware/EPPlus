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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Rounds a supplied number to a specified number of decimal places, and then converts this into text")]
    internal class Fixed : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToDecimal(arguments, 0);
            var nDecimals = 2;
            var noCommas = false;
            if (arguments.Count() > 1)
            {
                nDecimals = ArgToInt(arguments, 1);
            }
            if (arguments.Count() > 2)
            {
                noCommas = ArgToBool(arguments, 2);
            }
            var format = (noCommas ? "F" : "N") + nDecimals.ToString(CultureInfo.InvariantCulture);
            if (nDecimals < 0)
            {
                number = number - (number % (System.Math.Pow(10, nDecimals * -1)));
                number = System.Math.Floor(number);
                format = noCommas ? "F0" : "N0";
            }
            var retVal = number.ToString(format);
            return CreateResult(retVal, DataType.String);
        }
    }
}
