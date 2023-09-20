/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Financial,
       EPPlusVersion = "5.5",
       Description = "Calculates the Macauley modified duration for a security with an assumed par value of $100")]
    internal class Mduration : Duration
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var durationResult = base.Execute(arguments, context);
            if (durationResult.DataType == DataType.ExcelError) return durationResult;
            var dur = durationResult.ResultNumeric;
            var yield = ArgToDecimal(arguments, 3, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var frequency = ArgToDecimal(arguments, 4, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var result = dur / (1d + (yield / frequency));
            return CreateResult(result, DataType.Decimal);
        }
    }
}
