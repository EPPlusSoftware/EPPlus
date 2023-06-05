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
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Rounds a number up or down, to a given number of digits",
        SupportsArrays = true)]
    internal class Round : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = arguments[0];
            if (arg1.Value == null) return CreateResult(0d, DataType.Decimal);
            var nDigits = ArgToInt(arguments, 1);
            var positivDigits = nDigits * -1;
            
            var number = ArgToDecimal(arguments, 0, context.Configuration.PrecisionAndRoundingStrategy);
            if (nDigits < 0)
            {
                return CreateResult(System.Math.Round(number / System.Math.Pow(10, positivDigits), 0, MidpointRounding.AwayFromZero) * System.Math.Pow(10, positivDigits), DataType.Integer);
            }
            return CreateResult(System.Math.Round(number, nDigits, MidpointRounding.AwayFromZero), DataType.Decimal);
        }
        public override int ArgumentMinLength => 2;
    }
}
