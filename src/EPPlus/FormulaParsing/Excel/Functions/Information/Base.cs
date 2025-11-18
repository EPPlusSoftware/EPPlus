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

        public const string Digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var numberVal = arguments[0];
            var radixVal = arguments[1];

            var retRange = GetReturnRange(numberVal, radixVal);
            for(var row = 0; row < retRange.Size.NumberOfRows; row++)
            {
                for(var col = 0; col < retRange.Size.NumberOfCols; col++)
                {
                    var number = GetArgValue(numberVal, row, col);
                    var radix = GetArgValue(radixVal, row, col);
                    var res = GetBaseValue(number, radix);
                    retRange.SetValue(row, col, res);
                }
            }
            return CreateDynamicArrayResult(retRange, DataType.ExcelRange);           
        }

        private object GetArgValue(FunctionArgument arg, int row, int col)
        {
            if (arg.IsExcelRange)
            {
                var r = arg.ValueAsRangeInfo;
                if(r.Size.NumberOfRows == 1)
                {
                    if(r.Size.NumberOfCols == 1)
                    {
                        return r.GetOffset(0, 0);
                    }
                    if(r.Size.NumberOfCols <= col)
                    {
                        return ErrorValues.NAError; // Dessa kan ju inte vara här?
                    }
                    return r.GetOffset(0, col); 
                }
                else
                {
                    if (r.Size.NumberOfCols == 1)
                    {
                        return r.GetOffset(row, 0); 
                    }
                    if (r.Size.NumberOfRows <= row)
                    {
                        return ErrorValues.NAError;
                    }
                    return r.GetOffset(row, col);
                }
            }
            else
            {
                return arg.ValueFirst;
            }
        }

        private InMemoryRange GetReturnRange(FunctionArgument numberVal, FunctionArgument radixVal)
        {
            int rows, cols;
            if (numberVal.IsExcelRange)
            {
                rows = numberVal.ValueAsRangeInfo.Size.NumberOfRows;
                cols = numberVal.ValueAsRangeInfo.Size.NumberOfCols;
            }
            else
            {
                rows = 1;
                cols = 1;
            }
            if (radixVal.IsExcelRange)
            {
                var radixRange = radixVal.ValueAsRangeInfo;
                if (rows < radixRange.Size.NumberOfRows)
                {
                    rows = radixRange.Size.NumberOfRows;
                }
                if (cols < radixRange.Size.NumberOfCols)
                {
                    cols = radixRange.Size.NumberOfCols;
                }
            }
            return new InMemoryRange(rows, (short)cols);                
        }

        private object GetBaseValue(object value, object radix)
        {
            var result = string.Empty;
            if (value is ExcelErrorValue)
            {
                return value;
            }
            if(radix is ExcelErrorValue)
            {
                return radix;
            }
            //if(value is null || radix is null)
            //{
            //    return result = "0";
            //}
            var numberDec = ConvertUtil.GetValueDouble(value, true, true, true);
            var radixDec = ConvertUtil.GetValueDouble(radix, true, true, true);
            if (double.IsNaN(numberDec) || double.IsNaN(radixDec))
            {
                return ErrorValues.ValueError;
            }
            else if (numberDec < 0 || numberDec > Math.Pow(2, 53) || radixDec < 2 || radixDec > 36) 
            {
                return ErrorValues.NumError;
            }
            numberDec = Math.Truncate(numberDec);
            radixDec = Math.Truncate(radixDec);
            int numberInt = (int)numberDec;
            int radixInt = (int)radixDec;   
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
