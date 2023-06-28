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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal class ImSub : ImFunctionBase
    {
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = ArgToString(arguments, 0);
            var arg2 = ArgToString(arguments, 1);

            GetComplexNumbers(arg1, out double real, out double imag, out string imaginarySuffix);
            GetComplexNumbers(arg2, out double real2, out double imag2, out string imaginarySuffix2);

            var realPart = (real - real2);
            var imagPart = (imag - imag2);
            var sign = (imagPart < 0) ? "-" : "+";
           
            var usedPrefixes = GetUniquePrefixes(imaginarySuffix, imaginarySuffix2);
            var imSuffix = string.Empty;
            if (usedPrefixes.Count > 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            else if(usedPrefixes.Count == 1)
            {
                imSuffix = usedPrefixes[0];
            }
            var result = CreateImaginaryString(realPart, imagPart, sign, imSuffix);
            return CreateResult(result, DataType.String);
        }
    }
}
