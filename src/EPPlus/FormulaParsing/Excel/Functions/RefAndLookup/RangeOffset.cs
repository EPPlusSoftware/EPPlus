using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal class RangeOffset : ExcelFunction
    {
        public ExcelDataProvider.IRangeInfo StartRange{ get; set; }

        public ExcelDataProvider.IRangeInfo EndRange { get; set; }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {

            if (StartRange == null || EndRange == null) return CreateResult(eErrorType.Value);
            
            //Build the address from the minimum row and column to the maximum row and column. StartRange and offsetRange are single cells.
            var fromRow = System.Math.Min(StartRange.Address._fromRow, EndRange.Address._fromRow);
            var toRow = System.Math.Max(StartRange.Address._toRow, EndRange.Address._toRow);
            var fromCol = System.Math.Min(StartRange.Address._fromCol, EndRange.Address._fromCol);
            var toCol = System.Math.Max(StartRange.Address._toCol, EndRange.Address._toCol);
            var rangeAddress = new EpplusExcelDataProvider.RangeInfo(StartRange.Worksheet, new ExcelAddressBase(fromRow, fromCol, toRow, toCol));
            return CreateResult(rangeAddress, DataType.Enumerable);
        }
    }
}
