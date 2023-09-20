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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Rounds a number towards zero, (i.e. rounds a positive number down and a negative number up), to a given number of digits",
        SupportsArrays = true)]
    internal class Rounddown : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments[0].Value == null) return CreateResult(0d, DataType.Decimal);
            var number = ArgToDecimal(arguments, 0, out ExcelErrorValue e1, context.Configuration.PrecisionAndRoundingStrategy);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var nDecimals = ArgToInt(arguments, 1);

            var nFactor = number < 0 ? -1 : 1;
            number *= nFactor;

            double result;
            if (nDecimals > 0)
            {
                result = RoundDownDecimalNumber(number, nDecimals);
            }
            else
            {
                result = (int)Math.Floor(number);
                result = result - (result % System.Math.Pow(10, (nDecimals*-1)));
            }
            return CreateResult(result * nFactor, DataType.Decimal);
        }

        private static double RoundDownDecimalNumber(double number, int nDecimals)
        {
            var integerPart = System.Math.Floor(number);
            var decimalPart = number - integerPart;
            decimalPart = System.Math.Pow(10d, nDecimals)*decimalPart;
            decimalPart = System.Math.Truncate(decimalPart)/System.Math.Pow(10d, nDecimals);
            var result = integerPart + decimalPart;
            return result;
        }
    }
}
