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
       Description = "")]
    internal abstract class ImFunctionBase : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        internal void GetComplexNumbers(object arg, out double real, out double imag, out string imaginarySuffix)
        {
            if (arg is string formula)
            {
                formula = formula.Trim();
                var positions = formula.IndexOfAny(new char[] { '+', '-' },1);
                if (positions >= 0)
                {
                    GetNumbersFromPosition(out real, out imag, out imaginarySuffix, formula, positions);
                }
                else
                {
                    real = double.NaN;
                    imag = double.NaN;
                    imaginarySuffix = string.Empty;
                }
            }
            else
            {
                real = ConvertUtil.GetValueDouble(arg);
                imag = 0;
                imaginarySuffix = string.Empty;
            }
        }

        private static void GetNumbersFromPosition(out double real, out double imag, out string imaginarySuffix, string formula, int position)
        {
            var realString = formula.Substring(0, position);
            var imagString = formula.Substring(position);
            if (ConvertUtil.TryParseNumericString(realString, out real) == false)
            {
                real = double.NaN;
            }

            if (imagString.EndsWith("i") ||
                imagString.EndsWith("j"))
            {
                if (imagString.Length > 1 && (imagString[1] == '+' || imagString[1] == '-'))
                {
                    var sign = imagString[0] == imagString[1] ? "+" : "-";
                    imagString = sign + imagString.Substring(2);
                }
                if (ConvertUtil.TryParseNumericString(imagString.Substring(0, imagString.Length - 1), out imag) == false)
                {
                    
                    if (imagString.Length > 1 && (imagString.Substring(1).Equals("i") || imagString.Substring(1).Equals("j")))
                    {
                        imag = 1;
                    }
                    else
                    {
                        imag = double.NaN;
                    }

                }
                imaginarySuffix = imagString.Substring(imagString.Length - 1);
            }
            else
            {
                imag = double.NaN;
                imaginarySuffix = string.Empty;
            }
        }
    }
}
