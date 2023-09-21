/*************************************************************************************************
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
        EPPlusVersion = "7.0",
        Description = "Returns the inverse of the F probability distribution. If p = F.DIST(x,...), then F.INV(p,...) = x. The F distribution can be used in an F-test that compares the degree of variability in two data sets. For example, you can analyze income distributions in the United States and Canada to determine whether the two countries have a similar degree of income diversity.")]
    internal class FDotInv : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 3;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var probability = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var degreeOfFreedom1 = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);

            var degreeOfFreedom2 = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);

            degreeOfFreedom1 = Math.Floor(degreeOfFreedom1);
            degreeOfFreedom2 = Math.Floor(degreeOfFreedom2);

            if (probability < 0 || probability > 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            if (degreeOfFreedom1 < 1 || degreeOfFreedom2 < 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }

            var result = degreeOfFreedom2 / (degreeOfFreedom1 * (1 / BetaHelper.IBetaInv(probability, degreeOfFreedom1 / 2, degreeOfFreedom2 / 2) - 1));
            return CreateResult(result, DataType.Decimal);
        }
    }
}
