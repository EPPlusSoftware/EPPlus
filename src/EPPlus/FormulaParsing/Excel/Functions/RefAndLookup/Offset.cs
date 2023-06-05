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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns a reference to a range of cells that is a specified number of rows and columns from an initial supplied range")]
    internal class Offset : LookupFunction
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {           
            var arg0 = arguments[0];
            var rowOffset = ArgToDecimal(arguments, 1);
            var colOffset = ArgToDecimal(arguments, 2);
            var startRange = ArgToAddress(arguments, 0);
            
            int width = 0, height = 0;
            if (arguments.Count > 3)
            {
                height = ArgToInt(arguments, 3);
                if (height == 0) return new CompileResult(eErrorType.Ref);
            }
            if (arguments.Count > 4)
            {
                width = ArgToInt(arguments, 4);
                if (width == 0) return new CompileResult(eErrorType.Ref);
            }
            var adr = arg0.Address;
            if(adr==null) return new CompileResult(eErrorType.Value);
            var fromRow = adr.FromRow + (int)rowOffset;
            var fromCol = adr.FromCol + (int)colOffset;
            var toRow = (height != 0 ? adr.FromRow + height - 1 : adr.ToRow) + (int)rowOffset;
            var toCol = (width != 0 ? adr.FromCol + width - 1 : adr.ToCol) + (int)colOffset;

            var newRange = context.ExcelDataProvider.GetRange(adr.WorksheetName, fromRow, fromCol, toRow, toCol);
            
            return CreateAddressResult(newRange, DataType.ExcelRange);
        }
        public override int ArgumentMinLength => 3;

        public override bool ReturnsReference => true;
        public override bool HasNormalArguments => false;
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            if(argumentIndex==0)
            {
                return FunctionParameterInformation.IgnoreAddress;
            }
            return FunctionParameterInformation.Normal;
        }
        public override bool IsVolatile => true; 
    }
}
