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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal class ImSum : ImFunctionBase
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var argumentStrings = arguments[0].Value.ToString().Split(',');

            GetComplexNumbers(argumentStrings[0], out double real, out double imag, out string imaginarySuffix);
            GetComplexNumbers(argumentStrings[1], out double real2, out double imag2, out string imaginarySuffix2);

            if (double.IsNaN(real) || double.IsNaN(imag))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var realPart = (real + real2);
            var imagPart = (imag + imag2);

            if (imagPart == 0)
            {
                return CreateResult(realPart.ToString(), DataType.String);
            }
            var sign = (imagPart < 0) ? "-" : "+";
            var result = string.Format("{0}{1}{2}{3}", realPart, sign, Math.Abs(imagPart), imaginarySuffix);
            return CreateResult(result, DataType.String);
        }
    }
}
