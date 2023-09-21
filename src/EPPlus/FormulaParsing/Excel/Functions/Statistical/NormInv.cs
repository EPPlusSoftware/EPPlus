﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/29/2021         EPPlus Software AB       Implemented function
 *************************************************************************************************/
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
            EPPlusVersion = "5.8",
            Description = "Calculates the inverse of the Cumulative Normal Distribution Function for a supplied value of x, and a supplied distribution mean & standard deviation.")]
    internal class NormInv : NormInvBase
    {
        public override int ArgumentMinLength => 3;
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;
        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1, 2);
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var probability = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var mean = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            
            var stdev = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);
            
            if (probability < 0 || probability > 1 || stdev <= 0)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var result = NormsInv(probability, mean, stdev);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
