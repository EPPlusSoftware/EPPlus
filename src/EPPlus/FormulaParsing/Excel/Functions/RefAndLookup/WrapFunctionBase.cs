/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  XX/XX/XXXX         EPPlus Software AB           EPPlus vX
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    /// <summary>
    /// Base class for the WRAPROWS and WRAPCOLS functions. Provides shared
    /// argument parsing, vector flattening and validation.
    /// </summary>
    internal abstract class WrapFunctionBase : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;

        /// <summary>
        /// If the function is allowed in a pivot table calculated field.
        /// </summary>
        public override bool IsAllowedInCalculatedPivotTableField => false;

        /// <summary>
        /// Parses and validates the common WRAPROWS/WRAPCOLS arguments and flattens
        /// the source vector. Returns null on success; otherwise a CompileResult
        /// holding the error to return.
        /// </summary>
        /// <param name="arguments">The function arguments.</param>
        /// <param name="wrapCount">Out: the validated wrap count.</param>
        /// <param name="padValue">Out: the pad value (defaults to #N/A).</param>
        /// <param name="items">Out: the flattened source items.</param>
        /// <param name="errResult">Out: if an error occurs during parsing</param>
        protected void ParseArguments(
            IList<FunctionArgument> arguments,
            out int wrapCount,
            out object padValue,
            out List<object> items,
            out CompileResult errResult)
        {
            wrapCount = 0;
            padValue = ExcelErrorValue.Create(eErrorType.NA);
            items = null;
            errResult = default;

            // wrap_count must be present and a positive integer
            wrapCount = ArgToInt(arguments, 1, out ExcelErrorValue wrapErr);
            if (wrapErr != null)
            {
                errResult = CompileResult.GetDynamicArrayResultError(wrapErr.Type);
                return;
            }
            if (wrapCount < 1)
            {
                errResult = CompileResult.GetDynamicArrayResultError(eErrorType.Num);
                return;
            }

            // Optional pad value
            if (arguments.Count > 2 && arguments[2].Value != null)
            {
                padValue = arguments[2].Value;
            }

            // Collect the source values as a flat list. Input must be a 1D vector
            // (single row or single column). A 2D range yields #VALUE!.
            var firstArg = arguments[0];
            if (firstArg.IsExcelRange)
            {
                var range = firstArg.ValueAsRangeInfo;
                var rows = range.Size.NumberOfRows;
                var cols = range.Size.NumberOfCols;
                if (rows > 1 && cols > 1)
                {
                    errResult = CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                    return;
                }
                items = FlattenVector(range, rows, cols);
            }
            else
            {
                // Scalar input behaves as a single-element vector.
                items = new List<object>();
                items.Add(firstArg.Value);
            }
        }

        /// <summary>
        /// Flattens a 1D range (single row or single column) into a list of values.
        /// </summary>
        private static List<object> FlattenVector(IRangeInfo range, int rows, int cols)
        {
            var result = new List<object>(rows * cols);
            if (cols == 1)
            {
                for (var r = 0; r < rows; r++)
                {
                    result.Add(range.GetOffset(r, 0));
                }
            }
            else
            {
                for (var c = 0; c < cols; c++)
                {
                    result.Add(range.GetOffset(0, c));
                }
            }
            return result;
        }
    }
}