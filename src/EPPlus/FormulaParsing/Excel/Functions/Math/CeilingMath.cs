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
        EPPlusVersion = "5.1",
        Description = "Rounds a number up to the nearest integer or to the nearest multiple of significance",
        IntroducedInExcelVersion = "2013")]
    internal class CeilingMath : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToDecimal(arguments, 0);
            var significance = (arguments.Count() > 1) ? ArgToDecimal(arguments, 1) : 1;
            var mode = (arguments.Count() > 2) ? ArgToDecimal(arguments, 2) : 0d;
            if (RoundingHelper.IsInvalidNumberAndSign(number, significance)) return CreateResult(eErrorType.Num);
            return CreateResult(RoundingHelper.Round(number, significance, mode != 0d ? RoundingHelper.Direction.Up : RoundingHelper.Direction.AlwaysUp), DataType.Decimal);
        }
    }
}
