﻿/*************************************************************************************************
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
        EPPlusVersion = "5.1",
        Description = "Rounds a number down, regardless of the sign of the number, to a multiple of significance",
        IntroducedInExcelVersion = "2013",
        SupportsArrays = true)]
    internal class FloorPrecise : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments[0].Value == null) return CreateResult(0d, DataType.Decimal);
            var number = ArgToDecimal(arguments, 0, out ExcelErrorValue e1, context.Configuration.PrecisionAndRoundingStrategy);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var significance = 1d;
            if (arguments.Count > 1)
            {
                significance = ArgToDecimal(arguments, 1, 1, out ExcelErrorValue e2);
                if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            }
                
            if (RoundingHelper.IsInvalidNumberAndSign(number, significance)) return CompileResult.GetErrorResult(eErrorType.Num);
            return CreateResult(RoundingHelper.Round(number, significance, RoundingHelper.Direction.AlwaysDown), DataType.Decimal);
        }
    }
}
