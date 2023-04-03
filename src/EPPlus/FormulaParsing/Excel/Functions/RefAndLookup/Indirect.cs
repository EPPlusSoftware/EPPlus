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
using System.Xml.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table;

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
            FormulaRangeAddress adr;
            IRangeInfo result;
            var arg = arguments.First();
            if(arg.DataType==DataType.ExcelError)
            {
                return CompileResultFactory.Create(arg.ValueAsExcelErrorValue);
            }
            else if (ExcelCellBase.IsValidAddress(address))
            {                
                adr = new FormulaRangeAddress(context, address);
                result = context.ExcelDataProvider.GetRange(adr);
            }
            else if (ExcelAddressBase.IsTableAddress(address))
            {
                adr = new FormulaTableAddress(context, address);
                result = context.ExcelDataProvider.GetRange(adr);
            }
            else
            {
                //Check for external Worksbook
                int extRef, wsIx;
                ExcelCellBase.SplitAddress(ref address, out extRef, out wsIx, context.Package);
                var n = context.ExcelDataProvider.GetName(extRef, wsIx, address);
                if (n != null && n.Value is IRangeInfo ri)
                {
                    result=ri;
                }
                else
                {
                    return CompileResult.GetErrorResult(eErrorType.Name);
                }
            }

            if(result.IsRef)
            {
                return CompileResult.GetErrorResult(eErrorType.Ref);
            }
            else if (result.IsEmpty)
            {
                return CompileResult.Empty;
            }
            //else if(!result.IsMulti)
            //{
            //    var cell = result.FirstOrDefault();
            //    var val = cell != null ? cell.Value : null;
            //    if (val == null) return CompileResult.Empty;
            //    return CompileResultFactory.Create(val, result.Address);
            //}
            return new AddressCompileResult(result, DataType.ExcelRange, result.Address);
        }


        public override bool ReturnsReference => true;
    }
}
