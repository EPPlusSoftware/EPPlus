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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",        
        Description = "Converts a vertical range/array to a horizontal or vice versa.",
        SupportsArrays = true)]
    internal class Transpose : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg0 = arguments[0];
            if(arg0.DataType!=DataType.ExcelRange)
            {
                return CreateDynamicArrayResult(arg0.Value, arg0.DataType);
            }
            
            var range = arg0.ValueAsRangeInfo;
            var newRange = new InMemoryRange(new RangeDefinition(0, 0, (short)range.Size.NumberOfRows, range.Size.NumberOfCols));
            for(var r=0;r<range.Size.NumberOfRows;r++)
            {
                for (var c = 0; c < range.Size.NumberOfCols; c++)
                {
                    newRange.SetValue(c,r,range.GetOffset(r,c));
                }
            }

            return CreateDynamicArrayResult(newRange, DataType.ExcelRange);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
