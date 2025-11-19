/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  13/11/2025         EPPlus Software AB           EPPlus v8
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;


namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Information,
    EPPlusVersion = "8",
    Description = "Converts a number into a text representation with the given base.",
    SupportsArrays = true)]
    internal class Base : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;
        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1);
        }

        public const string Digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var numberVal = arguments[0].Value ?? 0d;
            var radixVal = arguments[1].Value ?? 0d;

            var result = string.Empty;            
            var numberDec = ConvertUtil.GetValueDouble(numberVal, true, true, true);
            var radixDec = ConvertUtil.GetValueDouble(radixVal, true, true, true);
            if (double.IsNaN(numberDec) || double.IsNaN(radixDec))
            {
                return CreateResult(ErrorValues.ValueError, DataType.ExcelError);
            }
            else if (numberDec < 0 || numberDec > Math.Pow(2, 53) || radixDec < 2 || radixDec > 36)
            {
                return CreateResult(ErrorValues.NumError, DataType.ExcelError);
            }
            numberDec = Math.Truncate(numberDec);
            radixDec = Math.Truncate(radixDec);
            int numberInt = (int)numberDec;
            int radixInt = (int)radixDec;
            result = GetBaseValue(numberInt, radixInt);
            if (arguments.Count > 2)
            {
                var minlength = arguments[2].Value;
                var minlengthDec = ConvertUtil.GetValueDouble(minlength, true, true, true);
                if(minlengthDec < 0d)
                {
                    return CreateResult(ErrorValues.NumError, DataType.ExcelError);
                }
                minlengthDec = Math.Truncate(minlengthDec);
                if (result.Length < minlengthDec)
                {
                    for (var i = result.Length; i < (int)minlength; i++)
                    {
                        result = "0" + result; 
                    }
                }
            }
            return CreateResult(result, DataType.String);         
        }
           

        private string GetBaseValue(int numberInt, int radixInt)
        {
            var result = string.Empty;
            if (numberInt == 0 || radixInt == 0) { result = "0"; }
            while (numberInt > 0)
            {
                int remainder = numberInt % radixInt;
                result = Digits[remainder] + result;
                numberInt /= radixInt;
            }            
            return result;
        }
    }
}
