using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal class RangeOffset : ExcelFunction
    {
        public string StartRange { get; set; }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            // EXAMPLE: A1:OFFSET(B2, 2, 0)

            var func = new Offset();
            // this call is OFFSET(B2, 2, 0)
            var offsetRangeResult = func.Execute(arguments, context);
            // Here is the result from the OFFSET function
            var offsetRange = offsetRangeResult.Result as ExcelDataProvider.IRangeInfo;
            if (offsetRange == null) return CreateResult(eErrorType.Value);

            // A1 should be set as StartRange by the UnregognizedFunctionName pipline in the FunctionExpression.
            
            //Build the address from the minimum row and column to the maximum row and column. StartRange and offsetRange are single cells.
            var startAddress = new ExcelAddressBase(StartRange);
            var fromRow = System.Math.Min(startAddress._fromRow, offsetRange.Address._fromRow);
            var toRow = System.Math.Max(startAddress._toRow, offsetRange.Address._toRow);
            var fromCol = System.Math.Min(startAddress._fromCol, offsetRange.Address._fromCol);
            var toCol = System.Math.Max(startAddress._toCol, offsetRange.Address._toCol);
            var rangeAddress = new EpplusExcelDataProvider.RangeInfo(offsetRange.Worksheet, new ExcelAddressBase(fromRow, fromCol, toRow, toCol));
            return CreateResult(rangeAddress, DataType.Enumerable);
            //throw new NotImplementedException();
        }
    }
}
