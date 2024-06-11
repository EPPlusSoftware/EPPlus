/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Allows filtering of a range or array data based on criteria.",
        SupportsArrays = true)]
    internal class FilterFunction : ExcelFunction
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = GetAsRangeInfo(arguments, 0);
            var arg2 = GetAsRangeInfo(arguments, 1);

            FunctionArgument arg3;
            if(arguments.Count > 2)
            {
                arg3 = arguments[2];
            }
            else
            {
                arg3 = null;
            }
            var s1 = arg1.Size;
            var s2 = arg2.Size;
            if (s1.NumberOfRows != s2.NumberOfRows && s1.NumberOfCols != s2.NumberOfCols)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }

            if(s1.NumberOfRows==s2.NumberOfRows)
            {
                return FilterOnRow(arg1, arg2, arg3);
            }
            else
            {
                return FilterOnColumn(arg1, arg2, arg3);
            }
        }
        public override int ArgumentMinLength => 2;
        private IRangeInfo GetAsRangeInfo(IList<FunctionArgument> arguments, int index)
        {
            var range = arguments[index].ValueAsRangeInfo;
            if (range == null)
            {
                var imr = new InMemoryRange(new RangeDefinition(1, 1));
                imr.SetValue(0, 0, arguments[index].Value);
                return imr;
            }
            return range;
        }


        private static CompileResult FilterOnRow(IRangeInfo arg1, IRangeInfo arg2, FunctionArgument emptyValue)
        {
            var s1 = arg1.Size;
            var s2 = arg2.Size;

            if (s2.NumberOfCols != 1)
            {
                return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
            }
            var filteredData = new List<List<object>>();
            var checkDimension = !arg2.IsInMemoryRange;
            var fr = arg2.Address.FromRow;
            var dfr = arg2.Dimension.FromRow;
            for (int r = 0; r < s2.NumberOfRows; r++)
            {
                if (checkDimension && fr + r > dfr)
                {
                    break;
                }
                var boolValue = ConvertUtil.GetValueDouble(arg2.GetOffset(r, 0)??false, false, true);
                if (double.IsNaN(boolValue))
                {
                     return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
                if (boolValue != 0)
                {
                    var row = new List<object>();

                    for (int c = 0; c < s1.NumberOfCols; c++)
                    {
                        row.Add(arg1.GetOffset(r, c));
                    }
                    filteredData.Add(row);
                }
            }
            if (filteredData.Count == 0)
            {
                if(emptyValue== null)
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Calc);
                }
                else
                {
                    return new DynamicArrayCompileResult(emptyValue.Value, emptyValue.DataType);
                }
            }
            return new DynamicArrayCompileResult(new InMemoryRange(filteredData), DataType.ExcelRange);
        }
        private static CompileResult FilterOnColumn(IRangeInfo arg1, IRangeInfo arg2, FunctionArgument emptyValue)
        {
            var s1 = arg1.Size;
            var s2 = arg2.Size;

            if (s2.NumberOfRows != 1)
            {
                return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
            }
            var filteredData = new List<List<object>>();
            //var nc = s2.NumberOfCols > arg1.Worksheet.Dimension.Columns ? s2.NumberOfCols : arg1.Worksheet.Dimension.Columns;
            var checkDimension = !arg2.IsInMemoryRange;
            var fc = arg2.Address.FromCol;
            var dfc = arg2.Dimension.FromCol;
            for (int c = 0; c < s2.NumberOfCols; c++)
            {
                if (checkDimension && fc + c > dfc)
                {
                    break;
                }
                var boolValue = ConvertUtil.GetValueDouble(arg2.GetOffset(0, c), false, true);
                if (double.IsNaN(boolValue))
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
                if (boolValue != 0)
                {
                    var row = new List<object>();

                    for (int r = 0; r < s1.NumberOfCols; r++)
                    {
                        row.Add(arg1.GetOffset(r, c));
                    }
                    filteredData.Add(row);
                }
            }
            if (filteredData.Count == 0)
            {
                if (emptyValue == null)
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Calc);
                }
                else
                {
                    return new DynamicArrayCompileResult(emptyValue.Value, emptyValue.DataType);
                }
            }
            return new DynamicArrayCompileResult(new InMemoryRange(filteredData), DataType.ExcelRange);
        }

        public override string NamespacePrefix
        {
            get
            {
                return "_xlfn._xlws.";
            }
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
