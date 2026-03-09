/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils.TypeConversion;
using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    internal static class RangeOperationsOperator
    {
        private const double DoublePrecision = 0.000000000000001d;

        private static CompileResult ApplyOperator(CompileResult l, CompileResult r, Operators op, out bool error, ParsingContext context)
        {
            error = false;
            if (!OperatorsEnumDict.Instance.ContainsKey(op))
            {
                error = true;
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var opImpl = OperatorsEnumDict.Instance[op];
            return opImpl.Apply(l, r, context);
        }

        internal static InMemoryRange Negate(IRangeInfo ri)
        {
            InMemoryRange imr;
            if (ri.IsInMemoryRange == false)
            {
                var physicalRows = RangeHelper.GetPhysicalRows(ri);
                var logicalRows = ri.Size.NumberOfRows;
                if (physicalRows < logicalRows)
                {
                    imr = new InMemoryRange(
                        new RangeDefinition(physicalRows, ri.Size.NumberOfCols),
                        logicalRows);
                }
                else
                {
                    imr = new InMemoryRange(ri.Size);
                }
            }
            else
            {
                imr = (InMemoryRange)ri;
            }

            int rows = imr.PhysicalRows;
            for (int c = 0; c < ri.Size.NumberOfCols; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    var d = ConvertUtil.GetValueDouble(ri.GetOffset(r, c), false, true);

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

            if (imr.HasVirtualRows)
            {
                // Compute the negation of an empty cell: -(0) = 0
                // Use the upstream default if one exists, otherwise null (= 0)
                var srcDefault = imr.VirtualDefaultValue;
                if (srcDefault == null)
                {
                    srcDefault = 0d;
                }
                var d = ConvertUtil.GetValueDouble(srcDefault, false, false);
                if (double.IsNaN(d))
                {
                    imr.VirtualDefaultValue = ErrorValues.ValueError;
                }
                else
                {
                    imr.VirtualDefaultValue = d == 0d ? 0d : -d;
                }
            }

            return imr;
        }

        private static InMemoryRange CreateRange(IRangeInfo l, IRangeInfo r, FormulaRangeAddress address)
        {
            var width = Math.Max(l.Size.NumberOfCols, r.Size.NumberOfCols);

            int logicalHeight = Math.Max(l.Size.NumberOfRows, r.Size.NumberOfRows);
            int physicalHeight = Math.Max(
                RangeHelper.GetPhysicalRows(l),
                RangeHelper.GetPhysicalRows(r));

            if (physicalHeight >= logicalHeight)
            {
                // No virtual rows needed - existing behavior
                var rangeDef = new RangeDefinition(logicalHeight, width);
                if (address != null)
                {
                    return new InMemoryRange(address, rangeDef);
                }
                else
                {
                    return new InMemoryRange(rangeDef);
                }
            }

            // Virtual range: small backing array, large logical size
            var physicalDef = new RangeDefinition(physicalHeight, width);
            if (address != null)
            {
                return new InMemoryRange(address, physicalDef, logicalHeight);
            }
            else
            {
                return new InMemoryRange(physicalDef, logicalHeight);
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

        private static void SetValue(Operators op, InMemoryRange resultRange, int row, int col, CompileResult leftVal, CompileResult rightVal, ParsingContext context)
        {
            var res = ApplyOperator(leftVal, rightVal, op, out bool error, context);
            var resultValue = res.ResultValue;
            if (!(resultValue is bool) && ConvertUtil.IsNumeric(resultValue) && res.ResultNumeric == 0d)
            {
                // avoid -0 results.
                resultValue = 0d;
            }
            SetValue(resultRange, row, col, resultValue, error);
        }

        private static bool ShouldUseSingleRow(RangeDefinition lSize, RangeDefinition rSize)
        {
            if ((lSize.NumberOfRows == 1 || rSize.NumberOfRows == 1) && lSize.NumberOfCols == rSize.NumberOfCols)
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

        private static bool SingleRowSingleCol(RangeDefinition lSize, RangeDefinition rSize)
        {
            return (lSize.NumberOfRows == 1 && rSize.NumberOfCols == 1) || (lSize.NumberOfCols == 1 && rSize.NumberOfRows == 1);
        }

        private static bool AddressIsNotAvailable(RangeDefinition lSize, RangeDefinition rSize, int row, int col)
        {
            if (row >= lSize.NumberOfRows || row >= rSize.NumberOfRows)
            {
                return true;
            }
            else if (col >= lSize.NumberOfCols || col >= rSize.NumberOfCols)
            {
                return true;
            }
            return false;
        }

        public static CompileResult Apply(CompileResult left, CompileResult right, Operators op, ParsingContext context)
        {
            if (left.DataType == DataType.ExcelRange && right.DataType != DataType.ExcelRange)
            {
                InMemoryRange resultRange = ApplySingleValueRight(left, right, op, context);
                return new DynamicArrayCompileResult(resultRange, DataType.ExcelRange, resultRange.Address, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
            }
            else if (left.DataType != DataType.ExcelRange && right.DataType == DataType.ExcelRange)
            {
                InMemoryRange resultRange = ApplySingleValueLeft(left, right, op, context);
                return new DynamicArrayCompileResult(resultRange, DataType.ExcelRange, resultRange.Address, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
            }
            if (left.DataType == DataType.ExcelRange && right.DataType == DataType.ExcelRange)
            {
                var interSectAddress = left.Address?.GetIntersectingRowOrColumns(right.Address);
                InMemoryRange resultRange = ApplyRanges(left, right, op, context, interSectAddress);
                return new DynamicArrayCompileResult(resultRange, DataType.ExcelRange, interSectAddress, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
            }
            return CompileResult.Empty;
        }

        private static object GetCellValue(IRangeInfo range, int rowOffset, int colOffset)
        {
            try
            {
                if (range.IsInMemoryRange || range.Address == null)
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
            catch
            {
                throw;
            }
        }

        private static void SetVirtualDefault(
            InMemoryRange resultRange,
            CompileResult nullLeft,
            CompileResult nullRight,
            Operators op,
            ParsingContext context)
        {
            if (!resultRange.HasVirtualRows) return;

            var res = ApplyOperator(nullLeft, nullRight, op, out bool error, context);
            if (error)
            {
                resultRange.VirtualDefaultValue = ExcelErrorValue.Create(eErrorType.Value);
            }
            else
            {
                var resultValue = res.ResultValue;
                if (!(resultValue is bool) && ConvertUtil.IsNumeric(resultValue) && res.ResultNumeric == 0d)
                {
                    resultValue = 0d;
                }
                resultRange.VirtualDefaultValue = resultValue;
            }
        }

        public static InMemoryRange ApplySingleValueRight(
            CompileResult left, CompileResult right, Operators op, ParsingContext context)
        {
            var lr = left.Result as IRangeInfo;
            if (lr == null && left.Result is FormulaRangeAddress fra)
            {
                lr = context.ExcelDataProvider.GetRange(fra);
            }
            else if (left.Address != null && left.Result is not InMemoryRange)
            {
                lr = context.ExcelDataProvider.GetRange(left.Address);
            }
            var resultRange = CreateRange(lr, InMemoryRange.Empty, lr.Address);
            for (var row = 0; row < resultRange.PhysicalRows; row++)
            {
                for (var col = 0; col < resultRange.Size.NumberOfCols; col++)
                {
                    var leftVal = GetCellValue(lr, row, col);
                    var lcr = CompileResultFactory.Create(leftVal);
                    SetValue(op, resultRange, row, col, lcr, right, context);
                }
            }

            // Compute default for virtual rows: null op scalar
            SetVirtualDefault(resultRange,
                CompileResultFactory.Create(null), right, op, context);

            return resultRange;
        }

        public static InMemoryRange ApplySingleValueLeft(
            CompileResult left, CompileResult right, Operators op, ParsingContext context)
        {
            var rr = right.Result as IRangeInfo;
            if (rr == null && right.Result is FormulaRangeAddress fra)
            {
                rr = context.ExcelDataProvider.GetRange(fra);
            }
            else if (right.Address != null && right.Result is not InMemoryRange)
            {
                rr = context.ExcelDataProvider.GetRange(right.Address);
            }
            var resultRange = CreateRange(InMemoryRange.Empty, rr, rr.Address);
            for (var row = 0; row < resultRange.PhysicalRows; row++)
            {
                for (var col = 0; col < resultRange.Size.NumberOfCols; col++)
                {
                    var leftVal = left.Result;
                    var rightVal = GetCellValue(rr, row, col);
                    var rcr = CompileResultFactory.Create(rightVal);
                    SetValue(op, resultRange, row, col, left, rcr, context);
                }
            }

            // Compute default for virtual rows: scalar op null
            SetVirtualDefault(resultRange,
                left, CompileResultFactory.Create(null), op, context);

            return resultRange;
        }

        private static void SetVirtualDefaultForRanges(
              InMemoryRange resultRange,
              IRangeInfo lr,
              IRangeInfo rr,
              bool shouldUseSingleRow,
              bool shouldUseSingleCell,
              bool singleRowSingleCol,
              Operators op,
              ParsingContext context)
        {
            if (!resultRange.HasVirtualRows) return;

            CompileResult virtualLeft, virtualRight;
            if (shouldUseSingleRow)
            {
                if (lr.Size.NumberOfRows == 1)
                {
                    virtualLeft = CompileResultFactory.Create(GetCellValue(lr, 0, 0));
                    virtualRight = CompileResultFactory.Create(null);
                }
                else
                {
                    virtualLeft = CompileResultFactory.Create(null);
                    virtualRight = CompileResultFactory.Create(GetCellValue(rr, 0, 0));
                }
            }
            else if (shouldUseSingleCell)
            {
                if (lr.Size.NumberOfCols == 1 && lr.Size.NumberOfRows == 1)
                {
                    virtualLeft = CompileResultFactory.Create(GetCellValue(lr, 0, 0));
                    virtualRight = CompileResultFactory.Create(null);
                }
                else
                {
                    virtualLeft = CompileResultFactory.Create(null);
                    virtualRight = CompileResultFactory.Create(GetCellValue(rr, 0, 0));
                }
            }
            else if (singleRowSingleCol)
            {
                if (lr.Size.NumberOfRows == 1)
                {
                    virtualLeft = CompileResultFactory.Create(GetCellValue(lr, 0, 0));
                    virtualRight = CompileResultFactory.Create(null);
                }
                else
                {
                    virtualLeft = CompileResultFactory.Create(null);
                    virtualRight = CompileResultFactory.Create(GetCellValue(rr, 0, 0));
                }
            }
            else
            {
                virtualLeft = CompileResultFactory.Create(null);
                virtualRight = CompileResultFactory.Create(null);
            }
            SetVirtualDefault(resultRange, virtualLeft, virtualRight, op, context);
        }

        private static InMemoryRange ApplyRanges(CompileResult left, CompileResult right, Operators op, ParsingContext context, FormulaRangeAddress intersectAddress)
        {
            var lr = left.Result as IRangeInfo;
            var rr = right.Result as IRangeInfo;
            if (lr == null && left.Result is FormulaRangeAddress fral)
            {
                lr = new RangeInfo(fral);
            }
            if (rr == null && right.Result is FormulaRangeAddress frar)
            {
                rr = new RangeInfo(frar);
            }

            var resultRange = CreateRange(lr, rr, intersectAddress);
            var shouldUseSingleCol = ShouldUseSingleCol(lr.Size, rr.Size);
            var shouldUseSingleRow = ShouldUseSingleRow(lr.Size, rr.Size);
            var shouldUseSingleCell = ShouldUseSingleCell(lr.Size, rr.Size);
            var singleRowSingleCol = SingleRowSingleCol(lr.Size, rr.Size);
            for (var row = 0; row < resultRange.PhysicalRows; row++)
            {
                for (var col = 0; col < resultRange.Size.NumberOfCols; col++)
                {
                    if (shouldUseSingleRow)
                    {
                        if (lr.Size.NumberOfRows == 1)
                        {
                            var leftVal = GetCellValue(lr, 0, col);
                            var rightVal = GetCellValue(rr, row, col);
                            SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
                        }
                        else if (rr.Size.NumberOfRows == 1)
                        {
                            var leftVal = GetCellValue(lr, row, col);
                            var rightVal = GetCellValue(rr, 0, col);
                            SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
                        }
                    }
                    else if (shouldUseSingleCol)
                    {
                        if (lr.Size.NumberOfCols == 1)
                        {
                            var leftVal = GetCellValue(lr, row, 0);
                            var rightVal = GetCellValue(rr, row, col);
                            SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
                        }
                        else if (rr.Size.NumberOfCols == 1)
                        {
                            var leftVal = GetCellValue(lr, row, col);
                            var rightVal = GetCellValue(rr, row, 0);
                            SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
                        }
                    }
                    else if (shouldUseSingleCell)
                    {
                        if (lr.Size.NumberOfCols == 1 && lr.Size.NumberOfRows == 1)
                        {
                            var leftVal = GetCellValue(lr, 0, 0);
                            var rightVal = GetCellValue(rr, row, col);
                            SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
                        }
                        else
                        {
                            var leftVal = GetCellValue(lr, row, col);
                            var rightVal = GetCellValue(rr, 0, 0);
                            SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
                        }
                    }
                    else if (singleRowSingleCol)
                    {
                        if (lr.Size.NumberOfRows == 1)
                        {
                            var leftVal = GetCellValue(lr, 0, col);
                            var rightVal = GetCellValue(rr, row, 0);
                            SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
                        }
                        else
                        {
                            var leftVal = GetCellValue(lr, row, 0);
                            var rightVal = GetCellValue(rr, 0, col);
                            SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
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
                        SetValue(op, resultRange, row, col, CompileResultFactory.Create(leftVal), CompileResultFactory.Create(rightVal), context);
                    }
                }
            }

            SetVirtualDefaultForRanges(resultRange, lr, rr,
                shouldUseSingleRow, shouldUseSingleCell, singleRowSingleCol,
                op, context);

            return resultRange;
        }
    }
}