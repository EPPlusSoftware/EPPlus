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
        Description = "Rounds a number towards zero, (i.e. rounds a positive number down and a negative number up), to a multiple of significance",
        SupportsArrays = true)]
    internal class Floor : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments[0].Value == null || arguments[1].Value == null) return CreateResult(0d, DataType.Decimal);
            var number = ArgToDecimal(arguments, 0, out ExcelErrorValue e1, context.Configuration.PrecisionAndRoundingStrategy);
            if (e1 != null) return CreateResult(e1.Type);
            var significance = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if(e2 != null) return CompileResult.GetErrorResult(e2.Type);
            if (RoundingHelper.IsInvalidNumberAndSign(number, significance)) return CompileResult.GetErrorResult(eErrorType.Num);
            return CreateResult(RoundingHelper.Round(number, significance, RoundingHelper.Direction.Down), DataType.Decimal);
        }
    }
}
