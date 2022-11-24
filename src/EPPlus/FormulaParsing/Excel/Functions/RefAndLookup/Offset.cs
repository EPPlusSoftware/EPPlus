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
        Description = "Returns a reference to a range of cells that is a specified number of rows and columns from an initial supplied range")]
    internal class Offset : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 3);
            var arg0 = functionArguments.First();
            var rowOffset = ArgToInt(functionArguments, 1);
            var colOffset = ArgToInt(functionArguments, 2);
            //var r = arg0.ValueAsRangeInfo;

            var startRange = ArgToAddress(functionArguments, 0, context);
            
            int width = 0, height = 0;
            if (functionArguments.Length > 3)
            {
                height = ArgToInt(functionArguments, 3);
                if (height == 0) return new CompileResult(eErrorType.Ref);
            }
            if (functionArguments.Length > 4)
            {
                width = ArgToInt(functionArguments, 4);
                if (width == 0) return new CompileResult(eErrorType.Ref);
            }
            var ws = context.Scopes.Current.Address.WorksheetName;            
            //var r =context.ExcelDataProvider.GetRange(ws, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol, startRange);
            var adr = arg0.Address;

            var fromRow = adr.FromRow + rowOffset;
            var fromCol = adr.FromCol + colOffset;
            var toRow = (height != 0 ? adr.FromRow + height - 1 : adr.ToRow) + rowOffset;
            var toCol = (width != 0 ? adr.FromCol + width - 1 : adr.ToCol) + colOffset;

            var newRange = context.ExcelDataProvider.GetRange(adr.WorksheetName, fromRow, fromCol, toRow, toCol);
            
            return CreateResult(newRange, DataType.Enumerable);
        }
        public override bool ReturnsReference => true;
       
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            if(argumentIndex==0)
            {
                return FunctionParameterInformation.IgnoreAddress;
            }
            return FunctionParameterInformation.Normal;
        }
    }
}
