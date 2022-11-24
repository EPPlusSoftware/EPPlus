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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns a cell or range reference that is represented by a supplied text string")]
    internal class Indirect : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var address = ArgToAddress(arguments, 0);
            ExcelAddressBase adr;
            if (ExcelAddressBase.IsValidAddress(address) || ExcelAddressBase.IsTableAddress(address))
            {
                adr = new ExcelAddressBase(address);
            }
            else
            {                
                var n=context.ExcelDataProvider.GetName(context.Scopes.Current.Address.WorksheetName, address);
                if(n.Value is IRangeInfo ri)
                {
                    adr = ri.Address.ToExcelAddressBase();
                }
                else
                {
                    adr = new ExcelAddressBase(n.Formula);
                }
                address = adr.Address;
            }
            var ws = adr.WorkSheetName;
            if (string.IsNullOrEmpty(ws))
            {
                ws = context.Scopes.Current.Address.WorksheetName;
            }
            var result = context.ExcelDataProvider.GetRange(ws, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol, address);
            if (result.IsEmpty)
            {
                return CompileResult.Empty;
            }
            else if(!result.IsMulti)
            {
                var cell = result.FirstOrDefault();
                var val = cell != null ? cell.Value : null;
                if (val == null) return CompileResult.Empty;
                return CompileResultFactory.Create(val);
            }
            return new CompileResult(result, DataType.Enumerable);
        }
        public override bool ReturnsReference => true;
    }
}
