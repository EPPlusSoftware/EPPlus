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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.Engineering,
            EPPlusVersion = "5.1",
            Description = "Calculates the modified Bessel function Yn(x)")]
    public class ConvertFunction : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var number = ArgToDecimal(arguments, 0);
            var fromUnit = ArgToString(arguments, 1);
            var toUnit = ArgToString(arguments, 2);
            if (!Conversions.IsValidUnit(fromUnit)) return CreateResult(eErrorType.NA);
            if (!Conversions.IsValidUnit(toUnit)) return CreateResult(eErrorType.NA);
            var result = Conversions.Convert(number, fromUnit, toUnit);
            if(double.IsNaN(result))
            {
                return CreateResult(eErrorType.NA);
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
