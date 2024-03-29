﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
    SupportsArrays = true,
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the chi-squared distribution.")]

    internal class ChiSqDotDist : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 3;
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1, 2);
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count > 3) return CompileResult.GetErrorResult(eErrorType.Value);

            var x = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var degFreedom = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);

            degFreedom = Math.Floor(degFreedom);
            var cumulative = ArgToBool(arguments, 2);

            if (x<0 || degFreedom<1 || degFreedom >Math.Pow(10,10) ) return CompileResult.GetErrorResult(eErrorType.Num);

            var result = 0d;

            if (cumulative)
            {
                result = ChiSquareHelper.CumulativeDistribution(x, degFreedom);
            }
            else
            {
                result = ChiSquareHelper.PropbabilityDistribution(x, degFreedom);
            }
 
            return CreateResult(result, DataType.Decimal);
        }
    }
}
