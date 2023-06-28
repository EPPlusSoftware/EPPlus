﻿/*************************************************************************************************
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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal class ImTan : ImFunctionBase
    {


        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            GetComplexNumbers(arguments[0].Value, out double real, out double imag, out string imaginarySuffix);
            if (double.IsNaN(real) || double.IsNaN(imag))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var realPart = (Math.Sin(real*2)/(Math.Cos(real*2)+Math.Cosh(imag*2)));
            var imagPart = (Math.Sinh(imag * 2) / (Math.Cos(real * 2) + Math.Cosh(imag * 2)));

            var sign = (imagPart < 0) ? "-" : "+";
            realPart = RoundingHelper.RoundToSignificantFig(realPart, 19);
            imagPart = RoundingHelper.RoundToSignificantFig(imagPart, 15);
            var result = string.Format("{0}{1}{2}{3}",realPart, sign, imagPart, imaginarySuffix);
            if (realPart == 0)
            {
                result = string.Format("{0}{1}", imagPart, imaginarySuffix);
                return CreateResult(result, DataType.String);
            }
            else
            {
                return CreateResult(result, DataType.String);

            }

        }
    }
}