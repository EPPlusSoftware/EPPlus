/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  13/11/2025         EPPlus Software AB           EPPlus v8
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Information,
    EPPlusVersion = "4",
    Description = "Tests a supplied value and returns an integer relating to the supplied value's error type",
    SupportsArrays = true)]
    internal class Base : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public const string Digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var number = ArgToInt(arguments, 0, out ExcelErrorValue errora0); // Must be trunkated 
            var radix = ArgToInt(arguments, 1, out ExcelErrorValue errora1);
            ExcelErrorValue errora2 = null;
            var minLength = arguments.Count() < 2 ? ArgToInt(arguments, 2, out errora2) : 0;
            if (errora0 != null) { return CompileResult.GetErrorResult(errora0.Type); }
            if (errora1 != null) { return CompileResult.GetErrorResult(errora1.Type); }
            if (errora2 != null) { return CompileResult.GetErrorResult(errora2.Type); }
            if(number < 0 || number > Math.Pow(2, 53) || radix < 2 || radix > 36 || minLength < 0) { return CompileResult.GetErrorResult(eErrorType.Num); }
            
            var result = string.Empty;
            while (number > 0)
            {
                int remainder = number % radix;
                result = Digits[remainder] + result;
                number /= radix;
            }

            return CreateResult(result, DataType.String);
        }
    }
}
