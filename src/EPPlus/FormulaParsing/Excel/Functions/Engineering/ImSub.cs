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
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var argumentStrings = arguments[0].Value.ToString().Split(',');

            GetComplexNumbers(argumentStrings[0], out double real, out double imag, out string imaginarySuffix);
            GetComplexNumbers(argumentStrings[1], out double real2, out double imag2, out string imaginarySuffix2);

            var realPart = (real - real2);
            var imagPart = (imag - imag2);
            var sign = (imagPart < 0) ? "-" : "+";
            var result = string.Format("{0}{1}{2}{3}", realPart, sign, Math.Abs(imagPart), imaginarySuffix);
            var imSuffix = imaginarySuffix;

            if (!string.IsNullOrEmpty(imaginarySuffix) && !string.IsNullOrEmpty(imaginarySuffix2) && imaginarySuffix != imaginarySuffix2)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            if (string.IsNullOrEmpty(imSuffix))
            {
                imSuffix = imaginarySuffix2;
            }
            if (imagPart == 1 || imagPart == -1)
            {
             
                result = string.Format("{0}{1}{2}", realPart, sign, imSuffix);
                return CreateResult(result, DataType.String);
            }
            else if (imagPart == 0)
            {
                result = string.Format("{0}", realPart);
                return CreateResult(result, DataType.String);
            }
            else if (realPart == 0)
            {
                result = string.Format("{0}{1}", imagPart, imSuffix);
                return CreateResult(result, DataType.String);
            }
            else
            {
                return CreateResult(result, DataType.String);
            }
        }
    }
}
