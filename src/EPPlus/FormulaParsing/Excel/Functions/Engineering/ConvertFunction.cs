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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    /// <summary>
    /// Convert function
    /// </summary>
    [FunctionMetadata(
            Category = ExcelFunctionCategory.Engineering,
            EPPlusVersion = "5.1",
            Description = "Converts a number from one measurement system to another")]
    public class ConvertFunction : ExcelFunction
    {
        /// <summary>
        /// Minimum arguments
        /// </summary>
        public override int ArgumentMinLength => 3;
        /// <summary>
        /// Execute convert function
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var number = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            var fromUnit = ArgToString(arguments, 1);
            var toUnit = ArgToString(arguments, 2);
            if (!Conversions.IsValidUnit(fromUnit)) return CompileResult.GetErrorResult(eErrorType.NA);
            if (!Conversions.IsValidUnit(toUnit)) return CompileResult.GetErrorResult(eErrorType.NA);
            var result = Conversions.Convert(number, fromUnit, toUnit);
            if(double.IsNaN(result))
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
