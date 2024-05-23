/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Returns the array in a single row.",
        SupportsArrays = true)]
    internal class ToRow : ToRowColBase
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var firstArg = arguments[0];
            var ignore = 0;
            if(arguments.Count > 1 && arguments[1].Value != null) 
            {
                ignore = ArgToInt(arguments, 1, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                if(ignore < 0 || ignore > 4)
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
            }
            var scanByColumn = false;
            if (arguments.Count > 2 && arguments[2].Value != null)
            {
                scanByColumn = ArgToBool(arguments, 2);
            }
            if (firstArg.IsExcelRange)
            {
                var range = firstArg.ValueAsRangeInfo;
                var result = GetItemsFromRange(firstArg.ValueAsRangeInfo, ignore, scanByColumn);
                var resultRange = new InMemoryRange(new RangeDefinition(1, (short)result.Count));
                var col = 0;
                foreach (var val in result)
                {
                    resultRange.SetValue(0, col++, val);
                }
                return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);  
            }
            return CompileResultFactory.CreateDynamicArray(firstArg.Value);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
