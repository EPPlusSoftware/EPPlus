/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    /// <summary>
    /// Returns error function
    /// </summary>
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Engineering,
        EPPlusVersion = "5.2",
        Description = "Returns the error function integrated between two supplied limits")]
    public class Erf : ExcelFunction
    {
        /// <summary>
        /// Min arguments
        /// </summary>
        public override int ArgumentMinLength => 1;
        /// <summary>
        /// Execute Erf
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var lowerLimit = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            var upperLimit = default(double?);
            if(arguments.Count > 1)
            {
                upperLimit = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
                if (e2 != null) return CreateResult(e2.Type);
            }
            var retVal = !upperLimit.HasValue ? ErfHelper.Erf(lowerLimit) : ErfHelper.Erf(lowerLimit, upperLimit.Value); 
            return CreateResult(retVal, DataType.Decimal);
        }
    }
}
