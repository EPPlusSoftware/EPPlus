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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
     Category = ExcelFunctionCategory.Engineering,
     EPPlusVersion = "7.0",
     Description = "Returns a complex number in x + yi or x + yj text format raised to a power")]
    internal class ImPower : ImFunctionBase
    {

       public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var comp = ArgToString(arguments, 0);
            var power = ArgToDecimal(arguments, 1, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            if (double.IsNaN(power))
            {     
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            GetComplexNumbers(comp, out double real, out double imag, out string imaginarySuffix);

            if (double.IsNaN(real) || double.IsNaN(imag))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var r = Math.Sqrt(Math.Pow(real, 2) + Math.Pow(imag, 2));
            var angle = Math.Atan(imag / real);
            var realPart = Math.Pow(r, power) * Math.Cos(power * angle);
            var imagPart = Math.Pow(r, power) * Math.Sin(power * angle);
            var sign = (imagPart < 0) ? "-" : "+";
            var result = CreateImaginaryString(realPart, imagPart, sign, imaginarySuffix);
            return CreateResult(result, DataType.String);
        }
    }
}