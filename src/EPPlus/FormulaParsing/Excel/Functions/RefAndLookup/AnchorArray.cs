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
using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Returns the range of the dynamic array starting at the cell-address supplied")]
    internal class AnchorArray : ExcelFunction
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var address = arguments[0].Address;
            if(address != null && address.WorksheetIx >= 0 && address.IsSingleCell && address.ExternalReferenceIx < 0)  //Not supported in external files yet
            {
                if (address.WorksheetIx>=0)
                {
                    var ws = context.Package.Workbook.Worksheets[address.WorksheetIx];
                    var f = ws._formulas.GetValue(address.FromRow, address.FromCol);
                    if(f is int sfIx)
                    {
                        var sf = ws._sharedFormulas[sfIx];
                        var rangeAddress = new FormulaRangeAddress(context)
                        {
                            FromRow = sf.StartRow,
                            FromCol = sf.StartCol,
                            ToRow = sf.EndRow,
                            ToCol = sf.EndCol,
                        };
                        var ri = new RangeInfo(rangeAddress);
                        return new DynamicArrayCompileResult(ri, DataType.ExcelRange, rangeAddress);
                    }
                }
            }
            return CompileResult.GetDynamicArrayResultError(eErrorType.Ref);
        }
        public override int ArgumentMinLength => 1;
        public override string NamespacePrefix
        {
            get
            {
                return "_xlfn.";
            }
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}    
}
