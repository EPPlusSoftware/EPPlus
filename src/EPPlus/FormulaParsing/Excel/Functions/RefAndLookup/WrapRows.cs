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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "X",
        Description = "Wraps a row or column vector into a 2D array of the specified number of columns per row.",
        SupportsArrays = true)]
    internal class WrapRows : WrapFunctionBase
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            int wrapCount;
            object padValue;
            List<object> items;
            ParseArguments(arguments, out wrapCount, out padValue, out items, out CompileResult error);
            if (error != null) return error;

            // Output shape: ceil(n / wrap_count) rows, wrap_count columns.
            var itemCount = items.Count;
            var resultRows = (itemCount + wrapCount - 1) / wrapCount;
            var resultRange = new InMemoryRange(new RangeDefinition(resultRows, (short)wrapCount));

            // Fill row-major: row 0 left-to-right, then row 1, etc.
            for (var i = 0; i < resultRows * wrapCount; i++)
            {
                var row = i / wrapCount;
                var col = i % wrapCount;
                object value;
                if (i < itemCount)
                {
                    value = items[i];
                }
                else
                {
                    value = padValue;
                }
                resultRange.SetValue(row, col, value);
            }

            return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
        }
    }
}