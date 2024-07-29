﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/30/2022         EPPlus Software AB       EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    internal static class RangeOperationsOperator
    {
        private const double DoublePrecision = 0.000000000000001d;
        private static object ApplyOperator(double l, double r, Operators op, out bool error)
        {
            error = false;
            switch(op)
            {
                case Operators.Plus:
                    return l + r;
                case Operators.Minus:
                    return l - r;
                case Operators.Multiply:
                    return l * r;
                case Operators.Divide:
                    return l / r;
                case Operators.LessThan:
                    return l < r;
                case Operators.LessThanOrEqual:
                    return l <= r;
                case Operators.GreaterThan:
                    return l > r;
                case Operators.GreaterThanOrEqual:
                    return l >= r;
                case Operators.Equals:
                    return Math.Abs(l - r) < DoublePrecision;
                case Operators.NotEqualTo:
                    return Math.Abs(l - r) > DoublePrecision;
                case Operators.Exponentiation:
                    return Math.Pow(l, r);
                case Operators.Concat:
                    var lRounded = RoundingHelper.RoundToSignificantFig(l, 15);
                    var rRounded = RoundingHelper.RoundToSignificantFig(r, 15);
                    return string.Concat(lRounded.ToString(CultureInfo.CurrentCulture), rRounded.ToString(CultureInfo.CurrentCulture));
                default:
                    error = true;
                    return default;
            }
        }

        private static object ApplyOperator(string l, string r, Operators op, out bool error)
        {
            error = false;
            switch(op)
            {
                case Operators.Concat:
                    return string.Concat(l, r);
                case Operators.LessThan:
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == false)
                    {
                        return true;
                    }
                    if(l != null && r == null)
                    {
                        return false;
                    }
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == true)
                        return false;
                    return string.Compare(l, r, StringComparison.InvariantCultureIgnoreCase) < 0;
                case Operators.LessThanOrEqual:
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == false)
                    {
                        return true;
                    }
                    if (string.IsNullOrEmpty(l) == false && string.IsNullOrEmpty(r) == true)
                    {
                        return false;
                    }
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == true)
                        return true;
                    return string.Compare(l, r, StringComparison.InvariantCultureIgnoreCase) <= 0;
                case Operators.GreaterThan:
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == false)
                    {
                        return false;
                    }
                    if (string.IsNullOrEmpty(l) == false && string.IsNullOrEmpty(r) == true)
                    {
                        return true;
                    }
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == true)
                        return false;
                    return string.Compare(l, r, StringComparison.InvariantCultureIgnoreCase) > 0;
                case Operators.GreaterThanOrEqual:
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == false)
                    {
                        return false;
                    }
                    if (string.IsNullOrEmpty(l) == false && string.IsNullOrEmpty(r) == true)
                    {
                        return true;
                    }
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == true)
                        return true;
                    return string.Compare(l, r, StringComparison.InvariantCultureIgnoreCase) >= 0;
                case Operators.Equals:
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == false)
                    {
                        return false;
                    }
                    if (string.IsNullOrEmpty(l) == false && string.IsNullOrEmpty(r) == true)
                    {
                        return false;
                    }
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == true)
                        return true;
                    return string.Compare(l, r, StringComparison.InvariantCultureIgnoreCase) == 0;
                case Operators.NotEqualTo:
                    if (string.IsNullOrEmpty(l)==true && string.IsNullOrEmpty(r)==false)
                    {
                        return true;
                    }
                    if (string.IsNullOrEmpty(l) == false && string.IsNullOrEmpty(r) == true)
                    {
                        return true;
                    }
                    if (string.IsNullOrEmpty(l) == true && string.IsNullOrEmpty(r) == true)
                        return false;
                    return string.Compare(l, r, StringComparison.InvariantCultureIgnoreCase) != 0;
                default:
                    error = true;
                    return null;
            }
        }
        internal static InMemoryRange Negate(IRangeInfo ri)
        {
            InMemoryRange imr;
            if(ri.IsInMemoryRange==false)
            {
                imr = new InMemoryRange(ri.Size);
            }
            else
            {
                imr = (InMemoryRange)ri;
            }

            for (int c = 0; c < ri.Size.NumberOfCols; c++)
            {
                for (int r = 0; r < ri.Size.NumberOfRows; r++)
                {
                    var d = ConvertUtil.GetValueDouble(ri.GetOffset(r, c), true, true);

                    if (double.IsNaN(d))
                    {
                        imr.SetValue(r, c, ErrorValues.ValueError);
                    }
                    else
                    {
                        imr.SetValue(r, c, -d);
                    }
                }
            }
            return imr;
        }
        private static InMemoryRange CreateRange(IRangeInfo l, IRangeInfo r, FormulaRangeAddress address)
        {
            var width = Math.Max(l.Size.NumberOfCols, r.Size.NumberOfCols);
            var height = Math.Max(l.Size.NumberOfRows, r.Size.NumberOfRows);
            var rangeDef = new RangeDefinition(height, width);
            if(address != null)
            {
                return new InMemoryRange(address, rangeDef);
            }
            else
            {
                return new InMemoryRange(rangeDef);
            }
        }

        private static void SetValue(InMemoryRange resultRange, int row, int col, object value, bool error)
        {
            if (!error)
            {
                resultRange.SetValue(row, col, value);
            }
            else
            {
                resultRange.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
            }
        }

        private static bool IsNumeric(object val)
        {
            return ConvertUtil.IsNumericOrDate(val, true, true);
        }

        private static void SetValue(Operators op, InMemoryRange resultRange, int row, int col, object leftVal, object rightVal)
        {
            if (IsNumeric(leftVal??0D) && IsNumeric(rightVal??0))
            {
                var l = ConvertUtil.GetValueDouble(leftVal, false, false, true);
                var r = ConvertUtil.GetValueDouble(rightVal, false, false, true);
                var result = ApplyOperator(l, r, op, out bool error);
                SetValue(resultRange, row, col, result, error);
            }
            else
            {
                var sResult = ApplyOperator(leftVal?.ToString(), rightVal?.ToString(), op, out bool error);
                SetValue(resultRange, row, col, sResult, error);
            }
        }

        private static bool ShouldUseSingleRow(RangeDefinition lSize, RangeDefinition rSize)
        {
            if((lSize.NumberOfRows == 1 || rSize.NumberOfRows == 1) && lSize.NumberOfCols == rSize.NumberOfCols)
            {
                return true;
            }
            return false;
        }

        private static bool ShouldUseSingleCol(RangeDefinition lSize, RangeDefinition rSize)
        {
            if ((lSize.NumberOfCols == 1 || rSize.NumberOfCols == 1) && lSize.NumberOfRows == rSize.NumberOfRows)
            {
                return true;
            }
            return false;
        }

        private static bool ShouldUseSingleCell(RangeDefinition lSize, RangeDefinition rSize)
        {
            return (lSize.NumberOfCols == 1 && lSize.NumberOfRows == 1) || (rSize.NumberOfCols == 1 && rSize.NumberOfRows == 1);
        }

        private static bool AddressIsNotAvailable(RangeDefinition lSize, RangeDefinition rSize, int row, int col)
        {
            if(row >= lSize.NumberOfRows || row >=rSize.NumberOfRows)
            {
                return true;
            }
            else if(col >= lSize.NumberOfCols || col >= rSize.NumberOfCols)
            {
                return true;
            }
            return false;
        }

        public static CompileResult Apply(CompileResult left, CompileResult right, Operators op, ParsingContext context)
        {
            if(left.DataType == DataType.ExcelRange && right.DataType != DataType.ExcelRange)
            {
                InMemoryRange resultRange = ApplySingleValueRight(left, right, op, context);
                return new AddressCompileResult(resultRange, DataType.ExcelRange, resultRange.Address);
            }
            else if(left.DataType != DataType.ExcelRange && right.DataType == DataType.ExcelRange)
            {
                InMemoryRange resultRange = ApplySingleValueLeft(left, right, op, context);
                return new AddressCompileResult(resultRange, DataType.ExcelRange, resultRange.Address);
            }
            if(left.DataType == DataType.ExcelRange && right.DataType == DataType.ExcelRange)
            {
                var interSectAddress = left.Address?.GetIntersectingRowOrColumns(right.Address);
                InMemoryRange resultRange = ApplyRanges(left, right, op, context, interSectAddress);
                return new AddressCompileResult(resultRange, DataType.ExcelRange, interSectAddress);
            }
            return CompileResult.Empty;
        }

        private static object GetCellValue(IRangeInfo range, int rowOffset, int colOffset)
        {
            if(range.IsInMemoryRange || range.Address == null)
            {
                return range.GetOffset(rowOffset, colOffset);
            }
            else
            {
                var col = range.Address.FromCol + colOffset;
                var row = range.Address.FromRow + rowOffset;
                return range.GetValue(row, col);
            }
        }

        public static InMemoryRange ApplySingleValueRight(CompileResult left, CompileResult right, Operators op, ParsingContext context)
        {
            var lr = left.Result as IRangeInfo;

            var resultRange = CreateRange(lr, InMemoryRange.Empty, lr.Address);
            for (var row = 0; row < resultRange.Size.NumberOfRows; row++)
            {
                for (var col = 0; col < resultRange.Size.NumberOfCols; col++)
                {
                    var leftVal = GetCellValue(lr, row, col);
                    SetValue(op, resultRange, row, col, leftVal, right.Result);
                }
            }
            return resultRange;
        }

        public static InMemoryRange ApplySingleValueLeft(CompileResult left, CompileResult right, Operators op, ParsingContext context)
        {
            var rr = right.Result as IRangeInfo;
            var resultRange = CreateRange(InMemoryRange.Empty, rr, rr.Address);
            for (var row = 0; row < resultRange.Size.NumberOfRows; row++)
            {
                for (var col = 0; col < resultRange.Size.NumberOfCols; col++)
                {
                    var leftVal = left.Result;
                    var rightVal = GetCellValue(rr, row, col);
                    SetValue(op, resultRange, row, col, leftVal, rightVal);
                }
            }
            return resultRange;
        }

        private static InMemoryRange ApplyRanges(CompileResult left, CompileResult right, Operators op, ParsingContext context, FormulaRangeAddress intersectAddress)
        {
            var lr = left.Result as IRangeInfo;
            var rr = right.Result as IRangeInfo;

            var resultRange = CreateRange(lr, rr, intersectAddress);
            var shouldUseSingleCol = ShouldUseSingleCol(lr.Size, rr.Size);
            var shouldUseSingleRow = ShouldUseSingleRow(lr.Size, rr.Size);
            var shouldUseSingleCell = ShouldUseSingleCell(lr.Size, rr.Size);
            for (var row = 0; row < resultRange.Size.NumberOfRows; row++)
            {
                for (var col = 0; col < resultRange.Size.NumberOfCols; col++)
                {
                    if (shouldUseSingleRow)
                    {
                        if (lr.Size.NumberOfRows == 1)
                        {
                            var leftVal = GetCellValue(lr, 0, col);
                            var rightVal = GetCellValue(rr, row, col);
                            SetValue(op, resultRange, row, col, leftVal, rightVal);
                        }
                        else if (rr.Size.NumberOfRows == 1)
                        {
                            var leftVal = GetCellValue(lr, row, col);
                            var rightVal = GetCellValue(rr, 0, col);
                            SetValue(op, resultRange, row, col, leftVal, rightVal);
                        }
                    }
                    else if (shouldUseSingleCol)
                    {
                        if (lr.Size.NumberOfCols == 1)
                        {
                            var leftVal = GetCellValue(lr, row, 0);
                            var rightVal = GetCellValue(rr, row, col);
                            SetValue(op, resultRange, row, col, leftVal, rightVal);
                        }
                        else if (rr.Size.NumberOfCols == 1)
                        {
                            var leftVal = GetCellValue(lr, row, col);
                            var rightVal = GetCellValue(rr, row, 0);
                            SetValue(op, resultRange, row, col, leftVal, rightVal);
                        }
                    }
                    else if (shouldUseSingleCell)
                    {
                        if (lr.Size.NumberOfCols == 1 && lr.Size.NumberOfRows == 1)
                        {
                            var leftVal = GetCellValue(lr, 0, 0);
                            var rightVal = GetCellValue(rr, row, col);
                            SetValue(op, resultRange, row, col, leftVal, rightVal);
                        }
                        else
                        {
                            var leftVal = GetCellValue(lr, row, col);
                            var rightVal = GetCellValue(rr, 0, 0);
                            SetValue(op, resultRange, row, col, leftVal, rightVal);
                        }
                    }
                    else if (AddressIsNotAvailable(lr.Size, rr.Size, row, col))
                    {
                        resultRange.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                    }
                    else
                    {
                        var leftVal = GetCellValue(lr, row, col);
                        var rightVal = GetCellValue(rr, row, col);
                        SetValue(op, resultRange, row, col, leftVal, rightVal);
                    }
                }
            }

            return resultRange;
        }
    }
}
