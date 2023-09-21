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
using System.Runtime.InteropServices;
using System.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "4",
        Description = "Tests a user-defined condition and returns one result if the condition is TRUE, and another result if the condition is FALSE")]
    internal class If : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.ElementAt(0).ValueIsExcelError)
            {
                return CompileResultFactory.Create(arguments.ElementAt(0).Value);
            }
            
            var arg0 = arguments[0].Value;
            var arg1 = arguments[1];
            var arg2 = arguments.Count < 3 ? new FunctionArgument(false,DataType.Boolean) : arguments[2];
            if (arg0 is IRangeInfo ri)
            {
                var arg1Type = GetType(arg1.Value);
                var arg2Type = GetType(arg2.Value);
                var range = new InMemoryRange(ri.Address, GetSizeForArray(ri, arg1, arg2));
                var isConditionSingleRow = ri.Size.NumberOfRows == 1;
                var isConditionSingleCol = ri.Size.NumberOfCols == 1;
                for (var row = 0; row < range.Size.NumberOfRows; row++)
                {
                    for (var col = 0; col < range.Size.NumberOfCols; col++)
                    {
                        var cellValue = ri.GetOffset(isConditionSingleRow ? 0 : row, isConditionSingleCol ? 0 : col);
                        var condition = ConvertUtil.GetValueBool(cellValue);
                        if (condition.HasValue)
                        {
                            object v = condition.Value ? GetArrayResult(arg1, arg1Type, row, col) : GetArrayResult(arg2, arg2Type, row, col);

                            range.SetValue(row, col, v);
                        }
                        else
                        {
                            if(cellValue is ExcelErrorValue error)
                            {
                                range.SetValue(row, col, error);
                            }
                            else
                            {
                                range.SetValue(row, col, ErrorValues.ValueError);
                            }
                        }
                    }
                }
                return new CompileResult(range, DataType.ExcelRange);
            }
            else
            {
                var condition = ConvertUtil.GetValueBool(arg0);
                if (condition.HasValue)
                {
                    if (arguments.Count < 3)
                    {
                        if (arg1.Address == null)
                        {
                            return condition.Value ? new CompileResult(arg1.Value, arg1.DataType) : CompileResultFactory.Create(false, null);
                        }
                        else
                        {
                            return condition.Value ? new AddressCompileResult(arg1.Value, arg1.DataType, arg1.Address) : CompileResultFactory.Create(false, null);
                        }
                    }
                    else
                    {
                        var secondStatement = arguments[2];
                        return condition.Value ?
                            arg1.Address == null ?
                                new CompileResult(arg1.Value, arg1.DataType) :
                                new AddressCompileResult(arg1.Value, arg1.DataType, arg1.Address) :
                            secondStatement.Address == null ?
                                new CompileResult(secondStatement.Value, secondStatement.DataType) :
                                new AddressCompileResult(secondStatement.Value, secondStatement.DataType, secondStatement.Address);
                    }
                }
                else
                {
                    if (arg0 is ExcelErrorValue error)
                    {
                        return CompileResult.GetErrorResult(error.Type);
                    }
                    else
                    {
                        return CompileResult.GetErrorResult(eErrorType.Value);
                    }
                }
            }
        }

        private RangeDefinition GetSizeForArray(IRangeInfo ri, FunctionArgument arg1, FunctionArgument arg2)
        {
            var rows = ri.Size.NumberOfRows;
            var cols = ri.Size.NumberOfCols;

            SetRowsColsFromSize(arg1, ref rows, ref cols);
            SetRowsColsFromSize(arg2, ref rows, ref cols);

            return new RangeDefinition(rows, cols); 
        }

        private static void SetRowsColsFromSize(FunctionArgument arg1, ref int rows, ref short cols)
        {
            if (arg1.DataType == DataType.ExcelRange)
            {
                var ri1 = arg1.Value as IRangeInfo;
                if (ri1 != null)
                {
                    rows = Math.Max(ri1.Size.NumberOfRows, rows);
                    cols = Math.Max(ri1.Size.NumberOfCols, cols);
                }
            }
        }

        public enum ArgumentType
        {
            Null,
            Number,
            Boolean,
            String,
            Range
        }
        private ArgumentType GetType(object value)
        {
            if(value==null)
            {
                return ArgumentType.Null;
            }
            var tc = Type.GetTypeCode(value.GetType());
            switch(tc)
            {
                case TypeCode.String:
                case TypeCode.Char:
                    return ArgumentType.String;
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Single: 
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.DateTime:
                    return ArgumentType.Number;
                case TypeCode.Boolean:
                    return ArgumentType.Boolean;
                case TypeCode.Empty:
                    return ArgumentType.Null;
                default:
                    if (value is IRangeInfo)
                        return ArgumentType.Range;
                    else
                        return ArgumentType.String;
            }
        }

        private object GetArrayResult(FunctionArgument arg, ArgumentType type, int row, int col)
        {
            if(type==ArgumentType.Range)
            {
                var r = arg.ValueAsRangeInfo;
                if(r.Size.NumberOfRows>row && r.Size.NumberOfCols>col)
                {
                    return r.GetOffset(row, col);
                }
                else if(r.Size.NumberOfRows > row && r.Size.NumberOfCols == 1)
                {
                    return r.GetOffset(row, 0);
                }
                else if (r.Size.NumberOfCols > col && r.Size.NumberOfRows == 1)
                {
                    return r.GetValue(0, col);
                }
                else if(r.Size.NumberOfCols == 1 && r.Size.NumberOfRows == 1)
                {
                    return r.GetValue(0, 0);
                }
                else
                {
                    return ExcelErrorValue.Create(eErrorType.NA);
                }
            }
            else
            {
                return arg.Value;
            }
        }

        public override bool ReturnsReference => true;
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex == 0)
            {
                return FunctionParameterInformation.Condition;
            }
            else if (argumentIndex == 1)
            {
                return FunctionParameterInformation.UseIfConditionIsTrue;
            }
            else
            {
                return FunctionParameterInformation.UseIfConditionIsFalse;
            }
        }));
    }
}
