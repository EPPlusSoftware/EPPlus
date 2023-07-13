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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Engineering,
       EPPlusVersion = "7.0",
       Description = "Returns the sum of two or more complex numbers in x + yi or x + yj text format.")]
    internal class ImSum : ImFunctionBase
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var realPart = 0d;
            var imagPart = 0d;
            var imSuffix = string.Empty;
            var args = new List<string>();

            foreach(var arg in arguments)
            {
                if (arg.IsExcelRange)
                {
                    foreach(var cell in arg.ValueAsRangeInfo)
                    {
                        if (cell.Value != null)
                        {
                            args.Add(cell.Value.ToString());
                        }
                    }
                }
                else if (arg.Value != null)
                {
                    args.Add(arg.Value.ToString());
                }
            }


            foreach (var argument in args )
            {
                GetComplexNumbers(argument, out double real, out double imag, out string imaginarySuffix);
                if (double.IsNaN(real) || double.IsNaN(imag))
                {
                    return CompileResult.GetErrorResult(eErrorType.Num);
                }
                if (!string.IsNullOrEmpty(imaginarySuffix))
                {
                    if (!string.IsNullOrEmpty(imSuffix) && imSuffix != imaginarySuffix)
                    {
                        return CompileResult.GetErrorResult(eErrorType.Value);
                    }
                    imSuffix = imaginarySuffix;
                }
                realPart += real;
                imagPart += imag;
            }

            if (imagPart == 0)
            {
                return CreateResult(realPart.ToString(), DataType.String);
            }
            var sign = (imagPart < 0) ? "-" : "+";
            var result = CreateImaginaryString(realPart, imagPart, sign, imSuffix);
            return CreateResult(result, DataType.String);
        }
    }
}
