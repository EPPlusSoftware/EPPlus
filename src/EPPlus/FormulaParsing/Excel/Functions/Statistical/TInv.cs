/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/07/2023         EPPlus Software AB         Implemented function
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
    Description = "Returns left tailed inverse of Students T-distribution")]
    internal class TInv : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {

            var probability = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var degreesOfFreedom = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if(e2 != null) return CompileResult.GetErrorResult(e2.Type);

            degreesOfFreedom = System.Math.Floor(degreesOfFreedom);

            if (probability <= 0 || probability > 1)
            {
                return CreateResult(eErrorType.Num);
            }

            if (degreesOfFreedom < 1)
            {
                return CreateResult(eErrorType.Num);
            }

            return CreateResult(StudenttHelper.InverseTFunc(probability, degreesOfFreedom), DataType.Decimal);
        }
    }
}
