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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Returns a random number between two given integers")]
    internal class RandBetween : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var low = ArgToDecimal(arguments, 0);
            var high = ArgToDecimal(arguments, 1);
            var rand = new Rand().Execute(new FunctionArgument[0], context).Result;
            var randPart = (CalulateDiff(high, low) * (double)rand) + 1;
            randPart = System.Math.Floor(randPart);
            return CreateResult(low + randPart, DataType.Integer);
        }

        private double CalulateDiff(double high, double low)
        {
            if (high > 0 && low < 0)
            {
                return high + low * - 1;
            }
            else if (high < 0 && low < 0)
            {
                return high * -1 - low * -1;
            }
            return high - low;
        }
        public override bool IsVolatile
        {
            get
            {
                return true;
            }
        }

    }
}
