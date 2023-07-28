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
            Category = ExcelFunctionCategory.Statistical,
            EPPlusVersion = "5.8",
            Description = "Calculates the inverse of the Cumulative Normal Distribution Function for a supplied value of x, and a supplied distribution mean & standard deviation.")]
    internal class NormsInv : NormInvBase
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 1;
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var probability = ArgToDecimal(arguments, 0);
            if (probability < 0 || probability > 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var result = NormsInv(probability, 0, 1);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
