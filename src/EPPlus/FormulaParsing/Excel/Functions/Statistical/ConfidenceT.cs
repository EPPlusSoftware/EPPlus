/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
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
        EPPlusVersion = "5.5",
        IntroducedInExcelVersion = "2010",
        Description = "Returns the confidence interval for a population mean, using a Student's t distribution")]
    internal class ConfidenceT : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var alpha = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var sigma = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            
            var size = ArgToInt(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);

            if (alpha <= 0d || alpha >= 1d) return CompileResult.GetErrorResult(eErrorType.Num);
            if (sigma <= 0d) return CompileResult.GetErrorResult(eErrorType.Num);
            if (size < 1d) return CompileResult.GetErrorResult(eErrorType.Num);

            var result = System.Math.Abs(StudentInv(alpha / 2, size - 1) * sigma / System.Math.Sqrt(size));
            return CreateResult(result, DataType.Decimal);
        }


        private double StudentInv(double p, double dof)
        {
            var x = BetaHelper.IBetaInv(2 * System.Math.Min(p, 1 - p), 0.5 * dof, 0.5);
            x = System.Math.Sqrt(dof * (1 - x) / x);
            return (p > 0.5) ? x : -x;
        }
    }
}
