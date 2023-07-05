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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
     Category = ExcelFunctionCategory.Engineering,
     EPPlusVersion = "7.0",
     Description = "Returns the cotangent of a complex number in x+yi or x+yj text format.")]
    internal class ImCot : ImFunctionBase
    {

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            GetComplexNumbers(arguments[0].Value, out double real, out double imag, out string imaginarySuffix);
            if (double.IsNaN(real) || double.IsNaN(imag))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            if (real >= 135000000)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }

            var realPart = 0d;
            var imagPart = 0d;
            if (real == 0)
            {
                realPart = 0;
                imagPart = -MathHelper.HCotan(imag);
            }
            else if (imag == 0)
            {
                realPart = MathHelper.Cotan(real);
                imagPart = 0;
            }
            else
            {
                var realNumerator = Math.Sin(2 * real);
                var realDenominator = -1 * (Math.Cos(2 * real) - Math.Cosh(2 * imag));
                realPart = realNumerator / realDenominator;

                var imagNumerator = -Math.Pow(MathHelper.Cotan(real), 2) * MathHelper.HCotan(imag) - MathHelper.HCotan(imag);

                var imagDenominator = Math.Pow(MathHelper.Cotan(real), 2) + Math.Pow(MathHelper.HCotan(imag), 2);
                imagPart = imagNumerator / imagDenominator;
            }

            var sign = (imagPart < 0) ? "-" : "+";
            var result = CreateImaginaryString(realPart, imagPart, sign, imaginarySuffix);
            if (double.IsNaN(imagPart))
            {
                sign = (imag > 0) ? "-" : "";
                if (double.IsNaN(realPart) || realPart == 0)
                {
                    realPart = 0;
                    imagPart = 1;
                    result = string.Format("{0}{1}", sign, imaginarySuffix);
                    return CreateResult(result, DataType.String);
                }
                else
                {
                    imagPart = 1;
                    result = string.Format("{0:G15}{1}{2}", realPart, sign, imaginarySuffix);
                    return CreateResult(result, DataType.String);
                }
            }
            else
            {
                return CreateResult(result, DataType.String);
            }
        }
    }
}
