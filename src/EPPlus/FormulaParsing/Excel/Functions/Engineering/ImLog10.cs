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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
     Category = ExcelFunctionCategory.Engineering,
     EPPlusVersion = "7.0",
     Description = "Returns the common logarithm (base 10) of a complex number in x + yi or x + yj text format.")]
    internal class ImLog10 : ImFunctionBase
    {


        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            GetComplexNumbers(arguments[0].Value, out double real, out double imag, out string imaginarySuffix);
            if (double.IsNaN(real) || double.IsNaN(imag))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var realPart = Math.Log(Math.Sqrt(Math.Pow(real, 2) + Math.Pow(imag, 2))) / Math.Log(10);
            var imagPart = Math.Atan(imag / real) / Math.Log(10);
            var sign = (imagPart < 0) ? "-" : "+";
            var result = CreateImaginaryString(realPart, imagPart, sign, imaginarySuffix);
            return CreateResult(result, DataType.String);
        }
    }
}

