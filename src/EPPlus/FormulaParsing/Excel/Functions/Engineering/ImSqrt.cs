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
     Description = "Returns the square root of a complex number in x + yi or x + yj text format.")]
    internal class ImSqrt : ImFunctionBase
    {


        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            GetComplexNumbers(arguments[0].Value, out double real, out double imag, out string imaginarySuffix);
            if (double.IsNaN(real) || double.IsNaN(imag))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var r = Math.Sqrt(real*real+imag*imag);
            var angle = Math.Atan(imag/real);
            var realPart = (Math.Sqrt(r) * Math.Cos(angle/2));

            var imagPart = (Math.Sqrt(r) * Math.Sin(angle / 2));

            var sign = (imagPart < 0) ? "-" : "+";
            realPart = RoundingHelper.RoundToSignificantFig(realPart, 15);
            imagPart = RoundingHelper.RoundToSignificantFig(imagPart, 15);
            var result = string.Format("{0}{1}{2}{3}", realPart, sign, Math.Abs(imagPart), imaginarySuffix);
            if (imagPart==1)
            {
                result = string.Format("{0}{1}{2}", realPart, sign, imaginarySuffix);
                return CreateResult(result, DataType.String);
            }
            else if (imagPart == 0)
            {
                result = string.Format("{0}", realPart);
                return CreateResult(result, DataType.String);
            }
            else 
            {
                return CreateResult(result, DataType.String);
            }
         }
    }
}