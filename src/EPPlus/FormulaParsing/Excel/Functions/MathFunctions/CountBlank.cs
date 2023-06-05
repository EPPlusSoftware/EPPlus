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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of blank cells in a supplied range")]
    internal class CountBlank : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg = arguments[0];
            if(!arg.IsExcelRange && arg.Address==null)throw new InvalidOperationException("CountBlank only support ranges as arguments");
            var result = 0;
            IRangeInfo range;
            if(arg.IsExcelRange)
            {
                range = arg.ValueAsRangeInfo;
                result =  arg.ValueAsRangeInfo.GetNCells();
            }
            else
            {
                //var currentCellAdr = context.CurrentCell;
                //var worksheet = currentCellAdr.WorksheetName;
                //var address = context.AddressCache.Get(arg.ExcelAddressReferenceId);
                //var excelAddress = new ExcelAddressBase(address);
                //if(!string.IsNullOrEmpty(excelAddress.WorkSheetName))
                //{
                //    worksheet = excelAddress.WorkSheetName;
                //}
                range = context.ExcelDataProvider.GetRange(arg.Address);
                result = range.GetNCells();
            }
            foreach (var cell in range)
            {
                if (!(cell.Value == null || cell.Value.ToString() == string.Empty))
                {
                    result--;
                }
            }
            return CreateResult(result, DataType.Integer);
        }
    }
}
