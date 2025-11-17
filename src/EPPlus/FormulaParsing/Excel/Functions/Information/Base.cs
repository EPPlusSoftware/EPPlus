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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        public const string Digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var numberVal = arguments[0];
            var radixVal = arguments[1];

            InMemoryRange numberRange = null;
            InMemoryRange radixRange = null;
            if (numberVal.IsExcelRange)
            {
                var src =  numberVal.ValueAsRangeInfo;
                numberRange = new InMemoryRange(src.Size);

                for (var row = 0; row < src.Size.NumberOfRows; row++)
                {
                    for (var col = 0; col < src.Size.NumberOfCols; col++)
                    {
                        var val = numberRange.GetOffset(row, col);
                        double num;
                        var ret = new InMemoryRange(numberRange.Size);
                        if (double.TryParse(val?.ToString(), out num))
                        {
                            if (num != Math.Truncate(num))
                            {
                                num = Math.Truncate(num);
                            }
                        }
                        numberRange.SetValue(row, col, num);
                    }
                }
            }
            else
            {
                var number = ArgToInt(arguments, 0, out ExcelErrorValue errora0);  
                if (errora0 != null) { return CompileResult.GetErrorResult(errora0.Type); }
                if(number < 0 || number > Math.Pow(2, 53)) { return CompileResult.GetErrorResult(eErrorType.Num); }
            }

            if (radixVal.IsExcelRange)
            {
                var src = radixVal.ValueAsRangeInfo;
                radixRange = new InMemoryRange(src.Size);

                for (var row = 0; row < radixRange.Size.NumberOfRows; row++)
                {
                    for (var col = 0; col < radixRange.Size.NumberOfCols; col++)
                    {
                        var val = radixRange.GetOffset(row, col);
                        double num;
                        var retRadixRange = new InMemoryRange(radixRange.Size);

                        if (double.TryParse(val?.ToString(), out num))
                        {
                            if (num != Math.Truncate(num))
                            {
                                num = Math.Truncate(num);
                            }
                        }
                        retRadixRange.SetValue(row, col, num);
                    }
                }
            }
            else
            {
                var radix = ArgToInt(arguments, 1, out ExcelErrorValue errora1);
                if (errora1 != null) { return CompileResult.GetErrorResult(errora1.Type); }
                if(radix < 2 || radix > 36) { return CompileResult.GetErrorResult(eErrorType.Num); }
            }

            var size = new RangeDefinition(numberRange.Size.NumberOfRows, radixRange.Size.NumberOfCols);
            var retRange = new InMemoryRange(size);
            for (int r = 0; r < retRange.Size.NumberOfRows; r++)
            {
                for (int c = 0; c < retRange.Size.NumberOfCols; c++)
                {
                    // Check if the cell exists in numberRange and radixRange
                    bool hasNumber = r < numberRange.Size.NumberOfRows;
                    bool hasRadix = c < radixRange.Size.NumberOfCols;

                    if (!hasNumber || !hasRadix)
                    {
                        // According to your rule #3:
                        retRange.SetValue(r, c, ExcelErrorValue.Create(eErrorType.Num));
                        continue;
                    }

                    // Get values
                    object numObj = numberRange.GetValue(r, 0);

                    if (numObj == null || !double.TryParse(numObj.ToString(), out double number))
                    {
                        retRange.SetValue(r, c, ExcelErrorValue.Create(eErrorType.Num));
                        continue;
                    }
                    object radObj = radixRange.GetValue(0, c);

                    if (radObj == null || !double.TryParse(radObj.ToString(), out double radix))
                    {
                        retRange.SetValue(r, c, ExcelErrorValue.Create(eErrorType.Num));
                        continue;
                    }


                    int numInt = (int)number;
                    int radInt = (int)radix;

                    // Base conversion
                    string baseStr = GetBaseValue(numInt, radInt);

                    retRange.SetValue(r, c, baseStr);
                }
            }

            //for (var row = 0; row < numberRange.Size.NumberOfRows; row++)
            //{
            //    for(var col = 0; col< radixRange.Size.NumberOfCols; col++)
            //    {

            //    }
            //}

            ExcelErrorValue errora2 = null;
            var minLength = arguments.Count() < 2 ? ArgToInt(arguments, 2, out errora2) : 0;
            if (errora2 != null) { return CompileResult.GetErrorResult(errora2.Type); }

            if(minLength < 0) { return CompileResult.GetErrorResult(eErrorType.Num); }
            
            

            return CreateResult(result, DataType.String);
        }

        private string GetBaseValue(int value, int radix)
        {
            var result = string.Empty;
            while (value > 0)
            {
                int remainder = value % radix;
                result = Digits[remainder] + result;
                value /= radix;
            }
            return result;
        }
    }
}
