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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal class Offset : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 3);
            var startRange = ArgToAddress(functionArguments, 0, context);
            var rowOffset = ArgToInt(functionArguments, 1);
            var colOffset = ArgToInt(functionArguments, 2);
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
            var ws = context.Scopes.Current.Address.Worksheet;            
            var r =context.ExcelDataProvider.GetRange(ws,startRange);
            var adr = r.Address;

            var fromRow = adr._fromRow + rowOffset;
            var fromCol = adr._fromCol + colOffset;
            var toRow = (height != 0 ? adr._fromRow + height - 1 : adr._toRow) + rowOffset;
            var toCol = (width != 0 ? adr._fromCol + width - 1 : adr._toCol) + colOffset;

            var newRange = context.ExcelDataProvider.GetRange(adr.WorkSheet, fromRow, fromCol, toRow, toCol);
            
            return CreateResult(newRange, DataType.Enumerable);
        }
    }
}
